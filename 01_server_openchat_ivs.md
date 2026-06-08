# OpenChat（AWS IVS）IVS処理順 技術ドキュメント

## 機能概要

OpenChatは1人のパフォーマが配信し、複数の会員・ゲストが視聴する一方向ライブ配信機能。映像配信はAWS IVS（Amazon Interactive Video Service）リアルタイムストリーミングの「Stage」を使う。FlaxはStageの作成・トークン発行・削除をSDK経由で管理し、配信状態をメモリ上で保持する。

**主要コンポーネント:**

| クラス | 役割 |
|--------|------|
| `IvsRealtimeUtils` | IVS SDK呼び出しのラッパー（Stage作成/削除、トークン発行、一覧） |
| `OpenChatRoomManager` | 配信のシングルトン管理。Stage作成/削除とOpenChatRoomの登録/解除 |
| `OpenChatRoom` | 1配信の状態（パフォーマ・Owner・stageArn・視聴者セッション） |
| `OpenChatPerformerWsClient` | 配信パフォーマ用WSエンドポイント `/ws/openChatPerformerLogin` |
| `OpenChatViewerWsClient` | 視聴用WSエンドポイント `/ws/openChatViewerLogin` |
| `OpenChatWsAction` | WSメッセージのaction名定数（フロントと一致させる） |

**対象パッケージ:** `jp.maru.flax.ivs` / `jp.maru.flax.openchat` / `jp.maru.flax.ws`

## IVS SDK呼び出し（IvsRealtimeUtils）

リージョンは東京（`ap-northeast-1`）固定。

| メソッド | IVS API | 用途 | トークン有効期限 |
|----------|---------|------|------------------|
| `createStage(name)` | CreateStage | 配信用Stage作成＋配信トークン取得 | 配信(publish) 1分 |
| `createSubscribeToken(stageArn)` | CreateParticipantToken | 視聴トークン発行 | 視聴(subscribe) 1分 |
| `deleteStage(stageArn)` | DeleteStage | Stage削除 | - |
| `listStages()` | ListStages | 孤児Stage検出用の全件取得（ページング） | - |

トークン有効期限が1分と短いのは、漏洩時のFlax外からの不正利用を防ぐため。トークンはWS接続直後のIVS join時のみ使用し、join後の配信/視聴の継続には不要。

Stage名は `openchat-{performerCode}-{timestamp}` 形式。`cleanupOrphans` がperformerCodeとtimestampをこの名前から復元する。

## IVS処理順サマリ

```
[配信開始] performer WS接続
   onOpen → 認証 → createRoom()
              └─ IVS: createStage()  ──► stageArn + publishToken
   STAGE_READY返却 → クライアントがStageへ publish join

[視聴開始] viewer WS接続（配信中のみ）
   onOpen → getRoom() → addViewer()
              └─ IVS: createSubscribeToken(stageArn) ──► subscribeToken
   JOINED返却 → クライアントがStageへ subscribe join

[視聴終了] viewer 切断
   onClose → removeViewer()        ※IVS呼び出しなし

[配信終了] performer 切断
   onClose → closeRoom()
              ├─ closeAllViewers()  → 全視聴者へ END通知しWSクローズ
              └─ IVS: deleteStage(stageArn)

[孤児削除] cron 毎時 cleanupOrphans()
   IVS: listStages() → 管理外の openchat- Stage を IVS: deleteStage()
```

IVS APIを呼ぶのは **配信開始（createStage）・視聴開始（createSubscribeToken）・配信終了（deleteStage）・孤児削除（listStages/deleteStage）** の4箇所のみ。視聴終了ではIVSを呼ばない（Stageは配信者のリソースのため）。

## シーケンス詳細

### 1. 配信開始（OpenChatPerformerWsClient.onOpen）

| 順 | 処理 | IVS |
|----|------|-----|
| 1 | origin(CORS) / x-forwarded-forからIP取得 | - |
| 2 | メンテナンス中チェック | - |
| 3 | token認証（`WebPerformerTokenAuth`）→ IP許可 → パフォーマ状態 | - |
| 4 | 通常チャットで配信中でないか（`isInNormalChat`） | - |
| 5 | 既にOpenChat配信中でないか（`getRoom`） | - |
| 6 | `createRoom(performer)` 内で **`createStage()`** | ● |
| 7 | `OpenChatRoom`生成しMap登録、`publishToken`/`expirationTime`取得 | - |
| 8 | `STAGE_READY {stageArn, publishToken, expirationTime}` 返却 | - |
| 9 | `startPingPong()` | - |

`createRoom` は `_creating` セットで同一パフォーマの同時作成を排他する。Stage作成後の登録処理で失敗した場合は **`deleteStage()` でロールバック**し、作成済みStageを残さない。

### 2. 視聴開始（OpenChatViewerWsClient.onOpen）

| 順 | 処理 | IVS |
|----|------|-----|
| 1 | origin(CORS)チェック | - |
| 2 | `performerCode` 必須・数値チェック | - |
| 3 | `token` があれば会員判定（無ければゲスト、無効トークンもゲスト扱い） | - |
| 4 | `getRoom(performerCode)`。null なら `NOT_STREAMING` 返却しクローズ | - |
| 5 | `addViewer(session, memberCode)`。同一会員の重複は拒否、close済みは `NOT_STREAMING` | - |
| 6 | **`createSubscribeToken(stageArn)`** で視聴トークン発行 | ● |
| 7 | `JOINED {stageArn, subscribeToken, expirationTime, isMember}` 返却 | - |
| 8 | `startPingPong()` | - |

`memberCode=0` はゲスト（重複チェック対象外）。会員は同一memberCodeで二重視聴不可。

### 3. 視聴終了（OpenChatViewerWsClient.onClose）

- Logoutメッセージ受信時は確認を返して `session.close()`、実処理は `onClose` に委ねる。
- `onClose` で `removeViewer()` を呼び、視聴者セッションとmemberCodeを解除。
- **IVS呼び出しなし**。Stageは配信者のリソースなので視聴者の離脱では削除しない。

### 4. 配信終了（OpenChatPerformerWsClient.onClose）

- Logoutメッセージ受信時は確認を返して `session.close()`、実処理は `onClose` に委ねる。
- `onClose` で `closeRoom(performerCode)` を呼ぶ。処理順は **Mapから除去 → 視聴者通知 → Stage削除**。

| 順 | 処理 | IVS |
|----|------|-----|
| 1 | `_openChatRoomMap` から除去（以降の新規視聴を遮断） | - |
| 2 | Owner別リストから除去 | - |
| 3 | `room.closeAllViewers()` → 全視聴者へ `END` 通知しWSクローズ | - |
| 4 | **`deleteStage(stageArn)`** | ● |

### 5. 孤児Stage削除（OpenChatRoomManager.cleanupOrphans、cron毎時）

FlaxInitializerのcronで（`0 * * * *`）1時間毎に実行。サーバ再起動やクラッシュで削除し損ねたStageを回収する。

| 順 | 処理 | IVS |
|----|------|-----|
| 1 | 現在Map上のstageArn一覧と `_creating` のスナップショットを取得 | - |
| 2 | **`listStages()`** で全Stage取得 | ● |
| 3 | `openchat-` 接頭辞以外、Map登録済みはスキップ | - |
| 4 | `_creating` 中のperformerCode、作成から5分以内のStageはスキップ | - |
| 5 | 残りを **`deleteStage()`** で削除 | ● |

作成5分以内をスキップするのは、`createStage` 直後でまだMap登録前の極短いraceで誤削除しないため。

## トークンと接続の関係

| ロール | WSパラメータ | 返却トークン | IVS join時のcapability |
|--------|--------------|--------------|------------------------|
| パフォーマ | `?token=...`（パフォーマ認証トークン） | `publishToken` | PUBLISH + SUBSCRIBE |
| 会員視聴 | `?performerCode=...&token=...` | `subscribeToken` | SUBSCRIBE |
| ゲスト視聴 | `?performerCode=...`（tokenなし） | `subscribeToken` | SUBSCRIBE |

WSの `token` はFlaxの認証トークンで、IVSのトークンとは別物。WSで認証した後にIVSトークンを払い出す二段構成。

## 注意点

- **視聴終了でdeleteStageしない**: Stageの削除は配信者の `closeRoom` だけ。視聴者離脱で消すと配信が落ちる。
- **closeRoomの順序**: Map除去を最初に行い、teardown中の新規視聴を防ぐ。`closeAllViewers` は1度だけ実行されるよう `_closed` フラグで保護。
- **ロールバック**: `createRoom` 途中失敗時は `deleteStage` で必ず作成済みStageを削除し、孤児を残さない。
- **孤児削除のスコープ**: `listStages()` はAWSアカウント＋リージョン内の全Stageを返す。`openchat-` 接頭辞に環境識別子が無いため、同一AWSアカウント・リージョンを複数環境で共有すると、他環境の配信中Stageを孤児と誤判定して削除する恐れがある。環境ごとにアカウント/リージョンを分けるか、接頭辞に環境識別子を入れること。
- **ping/pongタイムアウト**: 30秒間隔・2回連続未応答で約2分。切断検知が遅れると、その間Stageや視聴者カウントが残る。

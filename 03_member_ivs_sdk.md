# OpenChat Member IVS SDK（視聴側）

会員・ゲスト視聴用クライアントの AWS IVS Web Broadcast SDK 利用内容。

- SDK: `amazon-ivs-web-broadcast` v1.33.0
- 実装: `src/api/ivsStageController.js`
- 利用元: `src/components/ViewerComponent.vue`
- 役割: パフォーマの映像を IVS Stage から **subscribe（視聴）** する（自身は配信しない）

## import するSDK要素

| 要素 | 用途 |
|------|------|
| `Stage` | Stageへの接続オブジェクト |
| `SubscribeType` | 購読種別（視聴側は `AUDIO_VIDEO`） |
| `StageEvents` | Stageイベント定数 |
| `StageConnectionState` | 接続状態定数 |
| `StreamType` | ストリーム種別（VIDEO/AUDIO判定） |

配信側と違い `LocalStageStream` は使わない（publishしないため）。

## 処理順（ViewerComponent）

```
mounted
  joinStage(subscribeToken, onState, onStreams, onLayers)  // subscribe-only join
  ─ 視聴中 ─
    onStreams(mediaStream)  → video要素の srcObject にセット（映像表示）
    onLayers(layers)        → simulcast層リスト更新（画質選択UI）
    setPreferredLayer(label)→ 視聴画質を選択
beforeUnmount
  leaveStage()              // Stage退出
```

## joinStage(subscribeToken, onConnectionState, onStreams, onLayers)

subscribe専用の strategy を定義して接続する。

| strategy | 値 | 意味 |
|----------|-----|------|
| `stageStreamsToPublish` | `[]` | 配信しない |
| `shouldPublishParticipant` | `false` | 視聴者 |
| `shouldSubscribeToParticipant` | `SubscribeType.AUDIO_VIDEO` | 音声+映像を購読 |
| `preferredLayerForStream` | 関数 | 選択中のsimulcast層を返す（未選択なら自動） |

`new Stage(subscribeToken, strategy)` を生成し、以下のイベントを購読してから `await _stage.join()`。

## 購読イベント詳細

| イベント | 処理 |
|----------|------|
| `STAGE_CONNECTION_STATE_CHANGED` | connecting/connected/disconnected をコールバック通知 |
| `STAGE_PARTICIPANT_STREAMS_ADDED` | パフォーマ映像到着。`MediaStream` を組み立て `onStreams` で通知。映像の `getLayers()` を `onLayers` で通知 |
| `STAGE_PARTICIPANT_STREAMS_REMOVED` | 配信停止。`onStreams(null)` / `onLayers([])` で通知 |
| `STAGE_STREAM_LAYERS_CHANGED` | simulcast層の変化。選択中の層が消えたら自動に戻し（`refreshStrategy`）、`onLayers` で最新リスト通知 |

### STREAMS_ADDED の映像組み立て

```js
const mediaStream = new MediaStream()
for (const s of streams) mediaStream.addTrack(s.mediaStreamTrack)
onStreams(mediaStream)   // ViewerComponentが video.srcObject にセット
```

自身（`participant.isLocal`）のストリームは無視する。

## setPreferredLayer(label)

視聴する simulcast 層（画質）を選択。`_preferredLayerLabel` を更新し `_stage.refreshStrategy()` で反映。`label` が空なら自動選択に戻る。`preferredLayerForStream` がこの値を返すことで購読層が切り替わる。

## leaveStage()

`_stage.leave()` で退出し、`_stage` とコールバック参照（`_onStreams`/`_onLayers`/`_preferredLayerLabel`）を null化。

## WebSocketとの関係

- `subscribeToken` はWSの `JOINED` メッセージで受け取る（`{stageArn, subscribeToken, expirationTime, isMember}`）。
- トークン有効期限は約1分。WS接続後すみやかに `joinStage` する必要がある。
- 配信が終了するとサーバから `OpenChatEnd` が届きWSがクローズ。`leaveStage()` でStage退出。
- ログアウト/切断時はWS側で `Logout` 送信→サーバが `removeViewer`。視聴側ではIVS Stageの削除は行わない（Stageは配信者のリソース）。

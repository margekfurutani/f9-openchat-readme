# OpenChat（AWS IVS）ドキュメント一式

OpenChatライブ配信のIVS関連ドキュメントをまとめたもの。

| ファイル | 対象 | 内容 |
|----------|------|------|
| `01_server_openchat_ivs.md` | Flaxサーバ | IVS処理順（createStage/createSubscribeToken/deleteStage/listStages）、配信開始〜終了〜孤児削除のフロー |
| `02_performer_ivs_sdk.md` | 配信クライアント（openchat-performer） | publish側のWeb Broadcast SDK利用（Stage join・画質・simulcast・ミュート） |
| `03_member_ivs_sdk.md` | 視聴クライアント（openchat-member） | subscribe側のWeb Broadcast SDK利用（Stage join・購読イベント・画質層選択） |

## 構成概要

```
パフォーマ ──publish──► IVS Stage ◄──subscribe── 会員/ゲスト
    ▲                      ▲                        ▲
 performer WS          Flaxサーバ                viewer WS
 (publishToken)     (Stage作成/削除)           (subscribeToken)
```

- 配信側SDK: `amazon-ivs-web-broadcast` v1.33.0（publish）
- 視聴側SDK: `amazon-ivs-web-broadcast` v1.33.0（subscribe）
- サーバSDK: AWS SDK for Java `ivsrealtime`（Stage/トークン管理、リージョン ap-northeast-1）

## 原本の場所

| ドキュメント | 原本パス |
|--------------|----------|
| サーバ | `flax/.claude/docs/09_openchat_ivs.md` |
| 配信クライアント | `openchat-performer/docs/ivs-sdk.md` |
| 視聴クライアント | `openchat-member/docs/ivs-sdk.md` |

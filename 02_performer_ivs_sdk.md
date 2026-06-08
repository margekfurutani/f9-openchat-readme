# OpenChat Performer IVS SDK（配信側）

配信パフォーマ用クライアントの AWS IVS Web Broadcast SDK 利用内容。

- SDK: `amazon-ivs-web-broadcast` v1.33.0
- 実装: `src/api/ivsStageController.js`
- 利用元: `src/components/BroadcastComponent.vue`
- 役割: ローカルのカメラ/マイク映像を IVS Stage へ **publish（配信）** する

## import するSDK要素

| 要素 | 用途 |
|------|------|
| `Stage` | Stageへの接続オブジェクト |
| `LocalStageStream` | publishする音声/映像トラックのラッパー |
| `SubscribeType` | 購読種別（配信側は `NONE`） |
| `StageEvents` | Stageイベント定数 |
| `StageConnectionState` | 接続状態定数 |

## 画質プリセット（QUALITY_PRESETS）

| key | 解像度 | frameRate | maxBitrate(kbps) |
|-----|--------|-----------|-------------------|
| low | 640×360 | 30 | 800 |
| mid | 854×480 | 30 | 1200 |
| high | 1280×720 | 30 | 2500 |

frameRateは最大60まで設定可能。

## 処理順（BroadcastComponent）

```
mounted
  openLocalMedia(quality)        // getUserMediaでカメラ/マイク取得 → preview表示
  joinStage(publishToken, quality, onState)  // IVS Stageへ publish join
  ─ 配信中 ─
    setAudioMuted / setVideoMuted   // 音声・映像の送出ON/OFF
    updateQuality(quality)          // 画質変更（再joinなし）
beforeUnmount
  leaveStage()                   // Stage退出 + ローカルトラック停止
```

## 各メソッド詳細

### openLocalMedia(qualityKey)

`navigator.mediaDevices.getUserMedia` でカメラ/マイクを取得（解像度・frameRateはプリセットの `ideal` 指定）。取得した `MediaStream` を preview に表示。既存ストリームがあればトラックを `stop()` してから取り直す。

### joinStage(publishToken, qualityKey, onConnectionState)

1. `_buildPublishStreams()` で `LocalStageStream` を生成。
   - 音声: `new LocalStageStream(audioTrack)`
   - 映像: `new LocalStageStream(videoTrack, { maxBitrate, maxFramerate, simulcast: { enabled: true } })`
   - **simulcastを有効化** → 視聴側が複数画質レイヤーから選択できる。
2. publish専用の strategy を定義。

| strategy | 値 | 意味 |
|----------|-----|------|
| `stageStreamsToPublish` | `_publishStreams` | 音声/映像を配信 |
| `shouldPublishParticipant` | `true` | 自身は配信者 |
| `shouldSubscribeToParticipant` | `SubscribeType.NONE` | 他者は購読しない（一方向配信） |

3. `new Stage(publishToken, strategy)` を生成し、`STAGE_CONNECTION_STATE_CHANGED` を購読して接続状態（connecting/connected/disconnected）をコールバック通知。
4. `await _stage.join()` で接続。

### updateQuality(qualityKey)

配信中の画質変更。**Stageは維持したまま**メディアを取り直し、`_buildPublishStreams()` でトラックを差し替え、`_stage.refreshStrategy()` で反映する。再joinしないため `publishToken` は不要。

### setAudioMuted(muted) / setVideoMuted(muted)

publish中のトラックの送出を `LocalStageStream.setMuted()` で切り替える。`_publishStreams[0]`=音声、`[1]`=映像。

### leaveStage()

`_stage.leave()` で退出し、ローカル `MediaStream` の全トラックを `stop()`（カメラ/マイク解放）。`_stage` と `_localStream` を null化。

## WebSocketとの関係

- `publishToken` はWSの `STAGE_READY` メッセージで受け取る（`{stageArn, publishToken, expirationTime}`）。
- トークン有効期限は約1分。WS接続後すみやかに `joinStage` する必要がある。
- ログアウト/切断時はWS側で `Logout` 送信→サーバが配信終了（`deleteStage`）。クライアントは `leaveStage()` でローカル解放。

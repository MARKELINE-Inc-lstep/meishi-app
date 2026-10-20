# 名刺登録アプリ（MVP）

スマホで名刺を撮る → 文字を読み取る → 確認して登録 → 会社の共有スプレッドシートに1行 → AI が提案余地商材と想定される課題を入れる。

- 画面：`docs/index.html`（GitHub Pages で公開）
- 裏側：`Code.gs`（会社の共有 Google アカウントの Apps Script。読み取り・登録・AI判定はすべてこちら）
- `index.html`：Apps Script 側の画面（Google にログインしたブラウザでは開けないため、通常は Pages の画面を使う）

キー・パスワードはこのリポジトリに置かない（Apps Script のスクリプトプロパティに入れる）。

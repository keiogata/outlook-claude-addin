# Outlook Claude AI アシスタント

OutlookのタスクペインにClaude AIを組み込んだアドイン。  
自然言語でメール・予定表を検索・操作できます。

---

## 構成ファイル

| ファイル | 役割 |
|---------|------|
| `manifest.xml` | Outlookアドイン定義（リボンボタン・タスクペインURL） |
| `taskpane.html` | チャットUIのHTML |
| `taskpane.js` | メインロジック（チャット・Office.js連携） |
| `claude.js` | Anthropic Claude APIラッパー |
| `auth.js` | Microsoft Graph認証（Phase 2 スタブ） |
| `graph.js` | Microsoft Graph APIラッパー（Phase 2/3 スタブ） |
| `styles.css` | Fluent UI準拠スタイル |

---

## セットアップ手順

### Phase 1: Claude APIチャット機能（現在実装済み）

#### 1. GitHubリポジトリ作成

```bash
# リポジトリ名: outlook-claude-addin
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/YOUR_USERNAME/outlook-claude-addin.git
git push -u origin main
```

#### 2. GitHub Pages有効化

1. GitHubリポジトリ → Settings → Pages
2. Source: `main` ブランチの `/ (root)` を選択
3. 公開URL: `https://YOUR_USERNAME.github.io/outlook-claude-addin/`

#### 3. manifest.xmlのURL更新

`manifest.xml` 内の `YOUR_USERNAME` を自分のGitHubユーザー名に置換:

```xml
<!-- 変更前 -->
<SourceLocation DefaultValue="https://YOUR_USERNAME.github.io/..."/>

<!-- 変更後（例） -->
<SourceLocation DefaultValue="https://j0142312.github.io/outlook-claude-addin/taskpane.html"/>
```

#### 4. Outlookにアドインを登録

1. Outlookデスクトップを開く
2. 「ホーム」タブ → 「アドイン」 → 「アドインの管理」
3. 「カスタムアドイン」 → 「ファイルから追加」
4. `manifest.xml` を選択
5. Outlookを再起動 → リボンに「Claude AI」ボタンが表示される

#### 5. Anthropic APIキーの設定

1. タスクペイン右上の ⚙️ をクリック
2. [Anthropic Console](https://console.anthropic.com/) で取得したAPIキーを入力
3. 「保存」をクリック

> **注意**: APIキーはブラウザのlocalStorageに保存されます。個人PCでの使用を前提としています。

---

### Phase 2: 予定表機能（未実装）

#### 事前準備: Azure ADアプリ登録

1. [Azure Portal](https://portal.azure.com) → Microsoft Entra ID → アプリの登録 → 新規登録
2. アプリ名: `Outlook Claude アドイン`
3. リダイレクトURI: `https://YOUR_USERNAME.github.io/outlook-claude-addin/auth.html`
4. 「APIのアクセス許可」に以下を追加:
   - `User.Read`
   - `Calendars.Read`
   - `Calendars.Read.Shared`
   - `Mail.Read`
   - `People.Read`
5. 「クライアントID（アプリケーションID）」をコピー

#### auth.js の設定

```javascript
const AUTH_CONFIG = {
  clientId: 'XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX', // ← コピーしたIDを設定
  ...
};
```

---

### Phase 3: メール機能（未実装）

Phase 2完了後、`graph.js` の `searchEmails`・`getEmailBody` を実装。

---

## ローカルテスト方法

```bash
# Node.js必要
npx office-addin-dev-server start --document "https://outlook.office.com"
```

または、ローカルHTTPSサーバーを起動してOutlookのサイドロードを使用。

---

## 使用モデル

| 項目 | 設定値 |
|-----|--------|
| モデル | `claude-opus-4-6` |
| max_tokens | 2048 |
| 会話履歴保持 | 直近10往復 |
| ストリーミング | 有効 |

---

## トラブルシューティング

| 症状 | 対処 |
|-----|------|
| 「APIへの接続に失敗しました」 | プロキシ設定を確認。社内環境では`HTTPS_PROXY`を設定 |
| 「APIキーが無効です」 | Anthropic ConsoleでAPIキーを再確認 |
| リボンに「Claude AI」が表示されない | Outlookを再起動。manifest.xmlのURLが正しいか確認 |
| タスクペインが真っ白 | GitHub PagesのデプロイURL・ファイルパスを確認 |

---

## 注意事項

- 本アドインは個人利用を前提としています
- APIキーをコードにハードコードしないこと
- 社内機密情報をAPIに送信する際は情報セキュリティポリシーを確認してください
- Claude APIは社内プロキシを経由しない直接通信が必要な場合があります

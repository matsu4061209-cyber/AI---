# Claude Code コマンド一覧

> 最終更新: 2026-06-16

---

## 目次

1. [セッション管理](#セッション管理)
2. [コンテキスト・メモリ管理](#コンテキストメモリ管理)
3. [モデル・設定](#モデル設定)
4. [権限管理](#権限管理)
5. [ツール・統合管理](#ツール統合管理)
6. [表示・UI設定](#表示ui設定)
7. [認証・情報・その他](#認証情報その他)
8. [スキルベースのコマンド](#スキルベースのコマンド)

---

## セッション管理

| コマンド | 説明 |
|----------|------|
| `/clear [name]` | 新しい会話を開始（前の会話は `/resume` で復帰可能） |
| `/resume [session]` | 過去の会話を再開 |
| `/branch [name]` | 現在の会話から分岐を作成 |
| `/fork <指示>` | フォークサブエージェントを生成（バックグラウンドで実行） |
| `/rename [name]` | セッションに名前を付ける |
| `/rewind` / `/checkpoint` / `/undo` | 会話・コードを前の状態に戻す |
| `/exit` / `/quit` | Claude Code を終了 |

---

## コンテキスト・メモリ管理

| コマンド | 説明 |
|----------|------|
| `/memory` | CLAUDE.md ファイルを編集・自動メモリを管理 |
| `/context [all]` | コンテキスト使用量を視覚化 |
| `/compact [instructions]` | 会話を要約してコンテキストを解放（長い会話に有効） |

---

## モデル・設定

| コマンド | 説明 |
|----------|------|
| `/model [model]` | 使用するAIモデルを切り替え |
| `/effort [level\|auto]` | 推論レベルを設定（low / medium / high / xhigh / max / ultracode） |
| `/fast [on\|off]` | ファストモード切り替え（Opusを高速化） |
| `/config` / `/settings` | テーマ・出力スタイル等の設定画面を開く |
| `/goal [condition\|clear]` | ゴールを設定（条件達成まで自動実行） |
| `/sandbox` | サンドボックスモードを切り替え |

---

## 権限管理

| コマンド | 説明 |
|----------|------|
| `/permissions` / `/allowed-tools` | ツール実行権限ルールを管理 |
| `/hooks` | フック設定を表示 |
| `/privacy-settings` | プライバシー設定を表示 |

---

## ツール・統合管理

| コマンド | 説明 |
|----------|------|
| `/mcp [reconnect\|enable\|disable]` | MCPサーバーの接続を管理 |
| `/agents` | サブエージェント設定を管理 |
| `/plugin [subcommand]` | プラグインを管理（list / install / enable / disable 等） |
| `/add-dir <path>` | 作業ディレクトリを追加 |
| `/cd <path>` | セッションを新しいディレクトリに移動 |
| `/ide` | IDE統合の管理・ステータス表示 |
| `/chrome` | Chrome連携設定を管理 |
| `/install-github-app` | GitHub Actions連携をセットアップ |
| `/install-slack-app` | Slack連携をインストール |

---

## 表示・UI設定

| コマンド | 説明 |
|----------|------|
| `/theme` | カラーテーマを変更 |
| `/tui [default\|fullscreen]` | ターミナルUIレンダラーを設定 |
| `/color [color\|default]` | プロンプトバーの色を設定 |
| `/keybindings` | キーバインド設定ファイルを開く |
| `/statusline` | ステータスラインを設定 |
| `/voice [hold\|tap\|off]` | 音声入力モードを設定 |
| `/diff` | コミットされていない変更をdiffビューアで表示 |
| `/copy [N]` | 最後の応答をクリップボードにコピー |
| `/export [filename]` | 会話をテキストファイルで出力 |
| `/scroll-speed` | スクロール速度を調整 |

---

## 認証・情報・その他

| コマンド | 説明 |
|----------|------|
| `/help` | ヘルプと利用可能なコマンドを表示 |
| `/doctor` | Claude Codeのインストール・設定を診断 |
| `/status` | バージョン・モデル・アカウント情報を表示 |
| `/release-notes` | チェンジログ（更新履歴）を表示 |
| `/powerup` | 機能の対話的なレッスンを表示 |
| `/usage` / `/cost` / `/stats` | セッションコスト・使用状況を表示 |
| `/login` | Anthropic アカウントにサインイン |
| `/logout` | サインアウト |
| `/feedback` / `/bug` / `/share` | フィードバック・バグ報告を送信 |
| `/upgrade` | プラン選択ページを開く |
| `/desktop` / `/app` | Claude Code Desktopアプリで続行 |
| `/teleport` / `/tp` | ウェブセッションをターミナルに移動 |
| `/remote-control` / `/rc` | claude.ai からリモートコントロールを有効化 |
| `/mobile` / `/ios` / `/android` | Claude モバイルアプリのQRコードを表示 |
| `/radio` | Claude FM ラジオを開く |
| `/stickers` | Claude Code ステッカーを注文 |

---

## スキルベースのコマンド

スキルはビルトインコマンドとは別に追加される拡張コマンドです。  
`.claude/` ディレクトリにカスタムスキルを置くことで独自コマンドも追加できます。

### コードレビュー・品質

| コマンド | 説明 |
|----------|------|
| `/code-review [level] [--fix] [--comment]` | コード差分をレビュー。`ultra` でクラウド深掘りレビュー、`--fix` で自動修正、`--comment` でPRコメント投稿 |
| `/simplify [target]` | コード整理・効率化の提案を実施（重複排除・アルゴリズム改善） |
| `/security-review` | セキュリティ脅威をレビュー |
| `/review [PR]` | プルリクエストをローカルレビュー |
| `/ultrareview [PR]` | クラウドベースの深掘りレビュー（`/code-review ultra` の別名） |

### 開発・実行

| コマンド | 説明 |
|----------|------|
| `/init` | プロジェクト初期化（CLAUDE.md 生成） |
| `/run` | アプリを起動・動作確認 |
| `/verify` | コード変更を実装・検証 |
| `/debug [description]` | デバッグログを有効化・トラブルシューティング |
| `/batch <instruction>` | 大規模変更を複数ワーカーで並列実行 |

### 自動化・スケジュール

| コマンド | 説明 |
|----------|------|
| `/loop [interval] [prompt]` | プロンプトを反復実行（定期実行または自動ペース） |
| `/schedule [description]` | ルーティン（定期実行タスク）を作成・管理 |
| `/background [prompt]` | セッションをバックグラウンド実行 |
| `/stop` | バックグラウンドセッションを停止 |
| `/autofix-pr [prompt]` | CI失敗時に自動修正（ウェブセッション） |
| `/goal [condition]` | 条件達成まで自動実行するゴールを設定 |

### 分析・情報

| コマンド | 説明 |
|----------|------|
| `/deep-research <question>` | ウェブ検索・マルチソース検証・レポート生成 |
| `/insights` | セッション分析レポート生成 |
| `/recap` | セッションの1行要約を生成 |
| `/tasks` / `/bashes` | バックグラウンド実行状況を確認 |
| `/workflows` | ワークフロー進捗を監視 |
| `/btw <question>` | 会話に記録しない簡単な質問 |

### 設定・カスタマイズ

| コマンド | 説明 |
|----------|------|
| `/update-config` | Claude Code設定（settings.json）を変更・フック設定など |
| `/keybindings-help` | キーバインドをカスタマイズ |
| `/fewer-permission-prompts` | 権限プロンプト削減（許可リスト自動生成） |
| `/claude-api` | Claude API リファレンス・マイグレーション |

### その他スキル

| コマンド | 説明 |
|----------|------|
| `/plan [description]` | プランモードに進入して実装計画を立案 |
| `/ultraplan <prompt>` | 計画をブラウザで確認・実行 |
| `/team-onboarding` | チームオンボーディングガイド生成 |
| `/web-setup` | GitHub 連携をセットアップ |
| `/setup-bedrock` | Amazon Bedrock 認証設定 |
| `/setup-vertex` | Google Vertex AI 認証設定 |
| `/design-sync` | Reactデザインシステムを claude.ai/design にアップロード |
| `/run-skill-generator` | `/run` と `/verify` の実装をアプリに合わせてカスタマイズ |
| `/reload-plugins [--force]` | プラグイン再読み込み |
| `/reload-skills` | スキル再読み込み |

---

## コマンド使用上のポイント

| ポイント | 詳細 |
|----------|------|
| 入力位置 | メッセージの**先頭**でのみコマンドとして認識される |
| 候補表示 | `/` を入力すると利用可能なコマンドがフィルタ表示される |
| MCPコマンド | MCPサーバーが公開するプロンプトは `/mcp__<サーバー名>__<プロンプト名>` 形式 |
| カスタム追加 | `.claude/` ディレクトリにスキルファイルを置くと独自コマンドを追加できる |

---

*参考: [Claude Code 公式ドキュメント](https://docs.anthropic.com/claude-code)*

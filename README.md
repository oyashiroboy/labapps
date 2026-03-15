# labapps

研究・教育用途を主対象とした文書支援ツール集です。
現在は Ruby on Rails から Python スクリプトを呼び出して処理します。

## Target

- 研究者、大学院生、教員
- 論文・スライド・技術文書を扱うチーム
- 文書整形を半自動化したい開発者

## What This Repository Provides

- `ppt_tool`: 画像や設定値から PowerPoint 用テーマを生成　毎回コピーしてパワポを作成する手間から逃れましょう
- `vba_tool`: Excel ファイルの VBA 関連処理補助　現在補修中です
- `abbr_tool`: Wordファイルから略語抽出と検索し、略語の再定義や再使用によるリジェクトを防ぐことができます
- `unit_tool`: 単位表記チェックと CSV エクスポート SI単位系に則っているか、提出前のWordファイルをアップロードすると確認できます

## Architecture

- Web: Rails 8
- Script engine: Python (`python_scripts/*.py`)
- Data store: SQLite（デフォルト構成）
- Deployment baseline: Docker / Kamal 設定同梱

## License

本リポジトリは `AGPL-3.0-or-later` で公開します。

- フルテキスト: `LICENSE`
- SPDX: `AGPL-3.0-or-later`
- ネットワーク越し提供時も AGPL 条件が適用されます

### AGPL Operational Notes

- サービス提供時は、利用者がソースコードへ到達できる導線を UI 上に用意してください
- 変更版を運用する場合は、その変更を含む対応ソースを提供してください

### Commercial Use

商用利用（有償・無償を問わず、営利サービス組み込みを含む）は、別途デュアルライセンス契約を前提とします。
導入相談はリポジトリ保有者へ連絡してください。

## Language Policy

- 正本: 日本語
- 英語: 参考訳

## Local Development

### Requirements

- Ruby 3.3+
- Bundler
- Python 3.10+

### Setup

```bash
bundle install
pip install -r requirements.txt
bin/rails db:prepare
```

### Run

```bash
bin/dev
```

## Secrets Policy

`config/master.key` は Git にコミットしないでください。

- ローカル開発: `config/master.key` をローカルにのみ配置
- 本番/CI: `RAILS_MASTER_KEY` を環境変数として注入
- Kamal: `.kamal/secrets` は生値を置かず、環境変数参照のみを使用

```bash
export RAILS_MASTER_KEY="..."
bin/rails server
```

## Deployment
- `config/deploy.yml` を実運用値へ更新
- `RAILS_MASTER_KEY` を秘密情報として注入
- 必要に応じて `registry`, `servers`, `volumes` を環境に合わせて変更

## Security Posture

- アップロードファイルのサイズ、拡張子、MIME を検証
- Python 実行はタイムアウト付き subprocess で実行
- 詳細エラーはログへ、画面は一般化メッセージで返却
- 入力 JSON や mode 値は Rails 側と Python 側の双方で検証

脆弱性報告は公開 Issue ではなく、リポジトリ保有者へ直接連絡してください。

## OSS Release Checklist

- `config/master.key` が追跡対象でない
- `.env*`、秘密鍵、証明書類が追跡対象でない
- `bin/brakeman` に High 以上がない
- `bin/bundler-audit` に Critical/High がない
- Git 履歴に過去の秘密情報が残っていない

## Futures

- Python 処理の段階的 WebAssembly 化
- ローカル実行可能な Python Wasm ランタイム（Wasmer など）の導入検証
- サーバー側 Python 依存の縮小と再現性向上
- UI からの AGPL Source 導線の標準化

## Large-Scale Deployment

AGPL の範囲で利用できます。
大規模導入時は、保守連携のため事前連絡を歓迎します。

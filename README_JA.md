<p align="center">
  <a href="README.md">🇨🇳 中文</a> &nbsp;|&nbsp;
  <a href="README_EN.md">🇺🇸 English</a> &nbsp;|&nbsp;
  <a href="README_JA.md">🇯🇵 日本語</a>
</p>

---

# パラダイムレポート自動生成システム

> 差分分析に基づく Excel テンプレート抽出・自動入力ツール

## 解決する課題

業務用レポート Excel はバージョン間で 90% が同一で、一部のフィールド（作成者、バージョン番号、日付など）のみ変化します。手動での検索・修正は非効率でミスも発生します。

**本ツール**: 過去の Excel を投入 → 変動フィールドを自動識別 → 今回の変数を入力 → 新レポートを自動作成。

## クイックスタート

```bash
pip install -r requirements.txt
```

```bash
# 1. 過去レポートの分析
python3 main.py analyze --input ./data/history_reports --output ./project_config

# 2. 変数の入力（対話モードも可）
python3 main.py ask

# 3. 新レポートの生成
python3 main.py generate -t ./project_config/template.xlsx -v ./data/current_values/current_values.json -o ./output/report.xlsx
```

## コマンド一覧

| コマンド | 機能 |
|----------|------|
| `analyze` | 過去 Excel を分析し変動フィールドを自動識別 |
| `generate` | テンプレート + 変数 → 新レポート（グローバルテキスト置換含む） |
| `validate` | レポートの整合性を検証 |
| `diff` | 新旧レポートの比較 |
| `ask` | 対話形式で変数を入力 |

## プロジェクト構成

```
├── main.py              # CLI エントリポイント
├── src/
│   ├── excel_reader.py      # Excel 読み取り
│   ├── diff_analyzer.py     # 差分分析
│   ├── variable_detector.py # 変数検出
│   ├── template_builder.py  # テンプレート生成
│   ├── report_generator.py  # レポート生成 (XML レベル操作)
│   ├── validator.py         # 検証
│   └── change_logger.py     # 変更ログ
├── data/
├── project_config/
├── tests/
└── requirements.txt
```

## 技術スタック

Python 3.10+ · XML 直接操作 · 元フォーマット保持 · 17 テスト合格

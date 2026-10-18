---
name: sales-report-pptx
description: |
  営業訪問前レポートのスライド原稿（JSON）から、統一デザインの pptx を生成する。スライド構成の規約（8枚）、原稿の形式、生成スクリプトをまとめている。
  使う場面：report-writer が1社分の調査メモと廃棄物分析をもとにスライド原稿を書き、pptx を生成・再生成するとき。
  使わない場面：企業の調査（prospect-research を使う）、廃棄物の推定（waste-estimation を使う）、既存の pptx の読み取りや編集、訪問前レポート以外の汎用的なスライド作成。
---

# sales-report-pptx

スライド原稿を JSON で書き、スクリプトで pptx に変換する Skill です。デザイン（配色・フォント・レイアウト）はスクリプト側で統一するため、執筆担当は内容に集中できます。

## ファイル構成

| ファイル | 内容 |
|---|---|
| `references/slide_structure.md` | 標準のスライド構成（8枚）、原稿の形式、量の目安 |
| `templates/slides_sample.json` | 原稿のサンプル（主な type の書き方の例） |
| `scripts/build_pptx.py` | 原稿を検証して pptx を生成するスクリプト |

## 手順

1. `references/slide_structure.md` を読み、標準構成と形式を確認する。迷ったら `templates/slides_sample.json` を参照する。
2. スライド原稿を `output/draft/{社名}.json` に書く（UTF-8）。
3. プロジェクトのルートで次のコマンドを実行する。

   ```bash
   python .claude/skills/sales-report-pptx/scripts/build_pptx.py output/draft/{社名}.json
   ```

   出力先を省略すると `output/訪問前レポート_{社名}_{YYYYMMDD}.pptx`（実行日）に保存される。
4. 「スライド原稿にエラーがあります」と出たら、表示された箇所を直して再実行する。
5. 「生成しました: {パス}（n枚）」と表示されたら完了。

## 前提

- Python 3 と python-pptx が必要（`pip install -r requirements.txt`）。
- フォントは Meiryo を指定している。

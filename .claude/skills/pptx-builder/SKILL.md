---
name: pptx-builder
description: |
  M&A候補比較のスライド原稿（slides.json）から、統一デザインの pptx を生成する。スライド構成の規約、slides.json の形式、生成スクリプトをまとめている。
  使う場面：report-writer が3社の調査結果をもとに比較資料のスライド原稿を書き、pptx を生成・再生成するとき。
  使わない場面：企業の調査やレポート作成（ma-research-checklist を使う）、既存の pptx の読み取りや編集、M&A 比較資料以外の汎用的なスライド作成。
---

# pptx-builder

スライド原稿を JSON で書き、スクリプトで pptx に変換する Skill です。デザイン（配色・フォント・レイアウト）はスクリプト側で統一するため、執筆担当は内容に集中できます。

## ファイル構成

| ファイル | 内容 |
|---|---|
| `references/slide_structure.md` | 標準のスライド構成（10枚）、slides.json の形式、量の目安 |
| `templates/slides_sample.json` | slides.json のサンプル（全 type の書き方の例） |
| `scripts/build_pptx.py` | slides.json を検証して pptx を生成するスクリプト |

## 手順

1. `references/slide_structure.md` を読み、標準構成と形式を確認する。迷ったら `templates/slides_sample.json` を参照する。
2. スライド原稿を `output/draft/slides.json` に書く（UTF-8）。
3. プロジェクトのルートで次のコマンドを実行する。

   ```bash
   python .claude/skills/pptx-builder/scripts/build_pptx.py output/draft/slides.json
   ```

   出力先を省略すると `output/MA候補比較_{YYYYMMDD}.pptx`（実行日）に保存される。
4. 「slides.json にエラーがあります」と出たら、表示された箇所を直して再実行する。
5. 「生成しました: {パス}（n枚）」と表示されたら完了。

## 前提

- Python 3 と python-pptx が必要（`pip install -r requirements.txt`）。
- フォントは Meiryo を指定している。

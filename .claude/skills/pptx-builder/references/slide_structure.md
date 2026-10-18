# スライド構成の規約と slides.json の形式

## 標準構成（10枚）

| # | type | タイトル | 内容 |
|---|---|---|---|
| 1 | `title` | M&A候補先 比較資料 | サブタイトルに「候補先3社の公開情報に基づく初期評価」、日付 |
| 2 | `bullets` | エグゼクティブサマリー | 推奨順位（1〜3位）と一言の理由、次のアクション |
| 3 | `two_column` | 買い手の前提と評価軸 | 左：買収の目的／右：評価軸の優先順位と避けたい条件 |
| 4〜6 | `company` | 候補1〜3：{社名} | 1社1枚。facts に基本情報・財務（確度つき）、points に強み・懸念・DD 事項 |
| 7 | `table` | 3社比較 | 行＝評価軸（買い手の優先順位順）、列＝3社。◎○△ と短い根拠 |
| 8 | `table` | シナジーとリスク | 列＝社名／想定シナジー／主なリスク |
| 9 | `two_column` | 推奨順位と次のアクション | 左：推奨順位と理由／右：DD で確認すべき事項・次のアクション |
| 10 | `sources` | 出典一覧 | 3社のレポートの出典を `[n] タイトル - URL（確認日）` 形式で |

- 社の並び順は推奨順位順にする。
- エグゼクティブサマリーと推奨順位のスライドには、`note` で「公開情報に基づく初期評価であり、最終判断は DD の結果を踏まえて行う」旨を入れる。

## slides.json の形式

全体は `{"title": 資料名, "date": "YYYY-MM-DD", "footer": 任意, "slides": [...]}`。
すべてのスライドで任意の `note`（スライド下部の注記）と `notes`（発表者ノート）を使える。

| type | 必須フィールド | 任意フィールド |
|---|---|---|
| `title` | `title` | `subtitle`, `date` |
| `bullets` | `title`, `bullets` | — |
| `table` | `title`, `columns`, `rows` | `col_widths`（列幅の比率の配列） |
| `two_column` | `title`, `left`, `right`（それぞれ `{"heading", "bullets"}`） | — |
| `company` | `title`, `facts`, `points` | `summary`（1行の会社概要）, `points_heading` |
| `sources` | `title`, `items` | — |

- `bullets` の各要素は文字列、または `{"text": 文字列, "sub": [文字列, ...]}`（サブ項目つき）。
- `company` の `facts` は `{"label", "value", "confidence"}` の配列。`confidence` は `確認済み`／`推定`／`不明` のいずれか（確度ごとに色分けして表示される）。
- `table` の各行の要素数は `columns` と同じにする。

## 量の目安（はみ出し防止）

- `bullets`：サブ項目を含めて 10 行以内
- `table`：8 行以内、1セル 30 文字以内
- `company`：facts は 8 項目以内、points はサブ項目を含めて 8 行以内
- `sources`：24 件以内（13 件以上は自動で2列表示）

全体の形式は `templates/slides_sample.json` を参照する。

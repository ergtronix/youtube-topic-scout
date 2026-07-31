# youtube-topic-scout

YouTube新ジャンル発掘のためのマルチエージェントシステム。「視聴者の関心度が高く」「短期間で登録者が積み上がる可能性があり」「自宅で調査・とりまとめが完結し、自動生成できる（現地取材・出演不要）」という条件に合うYouTubeジャンルを、複数候補を並列調査したうえで比較・ランキングし、企画比較スライド（pptx）として出力する。

スクール教材『Claude Codeでマルチエージェントシステムを作ろう』第11章の最終課題として作成。課題用の使い捨てではなく、実際のYouTubeチャンネル運営における定期的な企画調査に運用する前提で設計している。

---

## 機能

- 候補ジャンル未指定時は、PM（メインエージェント）が「自宅完結・自動生成に向くジャンル」を自動でブレインストーム
- 複数の候補ジャンルを`genre-researcher`エージェントで並列調査（視聴者関心度・短期急成長ポテンシャル・自宅完結適性）
- `content-strategist`エージェントが調査結果を統一フォーマット（`genre-scorecard` Skill）で比較・ランキング
- ランキング結果を`slide-builder`エージェント（`md-to-slides` Skill）で企画比較スライド（pptx）として自動生成

## 使用技術（このシステム自体が満たす課題要件）

| 要件 | 対応 |
|---|---|
| カスタムサブエージェント2つ以上 | `genre-researcher.md`・`content-strategist.md`（`slide-builder.md`は流用） |
| Skills自作1つ以上 | `genre-scorecard`（3軸評価フォーマット） |
| Agentツールによる並列処理 | 候補ジャンルごとに`genre-researcher`を並列起動（ファンアウト） |
| 最終出力（スライド/動画） | 企画比較スライド（pptx） |

---

## 技術スタック

| 項目 | 内容 |
|---|---|
| 言語・環境 | Node.js（PptxGenJS・Playwright）／Python（スライド画像生成） |

---

## セットアップ

```bash
git clone https://github.com/ergtronix/youtube-topic-scout.git
cd youtube-topic-scout
npm install
```

---

## 使い方

Claude Codeでこのプロジェクトを開き、以下のように依頼する。

```
新しいYouTubeジャンルを調査してほしい
```

候補ジャンルを指定したい場合：

```
「都市伝説解説」「心理学要約」「歴史解説」の3ジャンルを調査してほしい
```

成果物は `output/{実行日}/scorecard.md`（比較表）・`output/{実行日}/slides/genre-scout.pptx`（比較スライド）に生成される。

---

## 振り返り

`output/retrospective.md` 参照。

---

## ライセンス

Private

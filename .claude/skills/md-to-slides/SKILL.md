---
name: md-to-slides
description: |
  MarkdownファイルをPowerPointスライド（pptx）とスライド画像（PNG）に変換するSkillです。
  使う場面：
  - 市場調査レポートやMarkdownドキュメントをスライドに変換するとき
  - output/{調査対象名（英小文字）}/slides/ 配下にpptxを、output/{調査対象名（英小文字）}/images/ 配下にPNGを生成したいとき
  使わない場面：
  - Markdownの内容修正・加筆（先にコンテンツを完成させてから使う）
  - PDFへの変換（本Skillはpptx＋PNG専用）
version: "1.0.0"
---

## 概要

Markdownファイル（h1/h2/h3構造を持つ）を解析してPowerPointスライドとスライド画像を生成します。

## 前提条件

- Node.js v20以上がインストールされていること
- 入力Markdownファイルが存在すること

## 手順

1. 入力Markdownファイルのパスを確認する
2. 出力先パス（output/{調査対象名（英小文字）}/slides/{調査対象名（英小文字）}.pptx）を決定する
3. 出力先ディレクトリが存在しない場合は作成する（generate.mjsはpptx出力先ディレクトリを自動作成しないため、事前に存在しないと書き込み時にENOENTエラーになる）：
   `mkdir -p output/{調査対象名（英小文字）}/slides/`
4. generate.mjs を実行してpptxとスライドHTMLを生成する（ブラウザで高さを実測する2段階方式）：
   `node .claude/skills/md-to-slides/scripts/generate.mjs {入力ファイル} {出力ファイル}`
5. render-slides.py を実行してスライド画像を生成する：
   `python .claude/skills/md-to-slides/scripts/render-slides.py output/{調査対象名（英小文字）}/images/slides/ output/{調査対象名（英小文字）}/images/`
6. 生成されたpptxと output/{調査対象名（英小文字）}/images/slide-NN.png のパスを報告する

## スライド構成ルール

- h1（# ）→ 表紙スライド（タイトルを中央に大きく表示）
- h2（## ）→ セクション名として記録（h3が続く場合はタイトル接頭辞として使用、h2直後に本文が続く場合はそのままスライドのタイトルとして使用）
- h3（### ）→「h2セクション名 — h3見出し」形式のスライドを生成
- h4以降（#### 以降）→ h3と同様に新スライドを生成（見出しテキストをタイトルとして左上に表示）
- ナレーション原稿は動画作成時にスライド画像から生成する

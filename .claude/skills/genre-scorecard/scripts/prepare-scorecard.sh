#!/bin/bash
# genre-scorecard/scripts/prepare-scorecard.sh
# スコアカード作成の前処理スクリプト
# 使い方: bash .claude/skills/genre-scorecard/scripts/prepare-scorecard.sh <入力ディレクトリ> <出力ファイル>
#
# 役割: 入力ディレクトリ・テンプレートの存在確認と出力先ディレクトリの作成を行う
#       実際のスコアリング・整形処理はClaudeが担当する

set -euo pipefail

INPUT_DIR="$1"
OUTPUT_FILE="$2"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
TEMPLATE="$SCRIPT_DIR/../templates/scorecard-template.md"
CRITERIA="$SCRIPT_DIR/../references/scoring-criteria.md"

if [ ! -d "$INPUT_DIR" ]; then
  echo "エラー：入力ディレクトリが見つかりません: $INPUT_DIR" >&2
  exit 1
fi

if [ ! -f "$TEMPLATE" ]; then
  echo "エラー：テンプレートが見つかりません: $TEMPLATE" >&2
  exit 1
fi

if [ ! -f "$CRITERIA" ]; then
  echo "エラー：採点基準ファイルが見つかりません: $CRITERIA" >&2
  exit 1
fi

FILE_COUNT=$(find "$INPUT_DIR" -maxdepth 1 -name "*.md" | wc -l)
if [ "$FILE_COUNT" -eq 0 ]; then
  echo "エラー：入力ディレクトリにジャンル調査ファイル（.md）が見つかりません: $INPUT_DIR" >&2
  exit 1
fi

OUTPUT_DIR="$(dirname "$OUTPUT_FILE")"
mkdir -p "$OUTPUT_DIR"

echo "=========================================="
echo "genre-scorecard 前処理スクリプト"
echo "=========================================="
echo "入力ディレクトリ：$INPUT_DIR（調査ファイル数: $FILE_COUNT）"
echo "出力：$OUTPUT_FILE"
echo "テンプレート：$TEMPLATE"
echo "採点基準：$CRITERIA"
echo "準備完了。Claudeがスコアリングを開始します。"

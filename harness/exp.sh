#!/bin/zsh
# Run each model × setup on a case file, then grade all of them. Usage: ./exp.sh cases_dev.json outdir
cd "$(dirname "$0")"
CASES=$1; OUT=$2; mkdir -p $OUT
for m in qwen2.5-coder:7b qwen3:4b-instruct gemma3:4b; do
  for v in +nocheck "" +plan +dyn +compact; do
    spec="$m$v"
    node run.js $CASES "$spec" >> $OUT/run.log 2>&1
    f="results-${spec//[:\/+]/_}.json"; mv "$f" $OUT/
    echo "ran $spec"
  done
done
python3 -I grade.py ../llm_test_sheet.xlsx $OUT/results-*.json > $OUT/grade.txt 2>&1
echo graded

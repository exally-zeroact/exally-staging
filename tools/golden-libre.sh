#!/usr/bin/env bash
# golden-libre.sh - LibreOffice headless で「LibreOfficeの答え」を作る。
#
#  ★これは Excel の真値ではない。LibreOffice という別の版の答え。
#    レポートでも「LibreOffice」列として別に出す。これで緑になっても「Excel一致」とは呼ばない。
#
#  なぜ apt install が要るか: LibreOffice は ubuntu-24.04 のGitHubランナー画像から削除されている
#    (https://github.com/actions/runner-images/blob/main/images/ubuntu/Ubuntu2404-Readme.md に記載が無い)。
#  なぜ再計算の設定が要るか: LibreOffice は xlsx を開いた時に既定で数式を計算し直さない。
#    OOXMLRecalcMode=1(常に再計算) を先に入れておく。
set -euo pipefail
ROOT="$(cd "$(dirname "$0")/.." && pwd)"
OUT="$ROOT/tests/xlsx-harness/golden"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT

command -v soffice >/dev/null 2>&1 || { echo "soffice が無い。sudo apt-get install -y libreoffice-calc"; exit 1; }
VER="$(soffice --version 2>/dev/null | head -1 | awk '{print $2}')"
echo "LibreOffice $VER"

# 読み込み時に必ず再計算する設定を仕込む
PROFILE="$TMP/profile"
mkdir -p "$PROFILE/user"
cat > "$PROFILE/user/registrymodifications.xcu" <<'XCU'
<?xml version="1.0" encoding="UTF-8"?>
<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema">
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>1</value></prop></item>
 <item oor:path="/org.openoffice.Office.Calc/Formula/Load"><prop oor:name="ODFRecalcMode" oor:op="fuse"><value>1</value></prop></item>
</oor:items>
XCU

# ケースから .xlsx を組み立てる(SheetJSで作る。式だけ入れてキャッシュ値は入れない=必ず計算させる)
node "$ROOT/tests/xlsx-harness/build-libre-input.mjs" "$TMP/cases.xlsx"

# 再計算させて csv に落とす
soffice --headless -env:UserInstallation="file://$PROFILE" \
        --convert-to csv:"Text - txt - csv (StarCalc)":44,34,76,1,,0,false,true,true \
        --outdir "$TMP" "$TMP/cases.xlsx" >/dev/null

node "$ROOT/tests/xlsx-harness/collect-libre.mjs" "$TMP/cases.csv" "$VER" "$OUT"
echo "golden/libreoffice-*.json を書き出した(★Excelの真値ではない。参考列)"

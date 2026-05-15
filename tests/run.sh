#!/usr/bin/env bash
# Headless test runner. Exercises the CLI paths that don't need a TTY.
# Usage: bash tests/run.sh
#
# Exits non-zero on the first failure. Each check prints a one-line PASS/FAIL.

set -uo pipefail

REPO_ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
cd "$REPO_ROOT"

PASS=0
FAIL=0

pass() { echo "PASS  $1"; PASS=$((PASS+1)); }
fail() { echo "FAIL  $1"; FAIL=$((FAIL+1)); }

# Build once, suppress noise. Fail loudly on build errors.
if ! dotnet build ExcelConsole.csproj -nologo -v q >/tmp/qs-build.log 2>&1; then
  echo "build failed:"
  cat /tmp/qs-build.log
  exit 1
fi
pass "build clean"

RUN=(dotnet run --project ExcelConsole.csproj --no-build --)

# ── CLI flag checks ────────────────────────────────────────────────

if "${RUN[@]}" --version 2>/dev/null | grep -q '^QuickSheet '; then
  pass "--version prints version line"
else
  fail "--version output unexpected"
fi

HELP_OUT=$("${RUN[@]}" --help 2>/dev/null || true)
if echo "$HELP_OUT" | grep -q 'Cell prefixes' && \
   echo "$HELP_OUT" | grep -qF -- '--export-md' && \
   echo "$HELP_OUT" | grep -qF -- '--list-extensions'; then
  pass "--help covers cell prefixes + export-md + list-extensions"
else
  fail "--help missing expected sections"
fi

# --export-md to a file
TMP_MD=$(mktemp --suffix=.md)
"${RUN[@]}" tests/fixtures/simple.csv --export-md "$TMP_MD" >/dev/null 2>&1
if diff -q tests/fixtures/simple.expected.md "$TMP_MD" >/dev/null; then
  pass "--export-md matches expected fixture"
else
  fail "--export-md output differs:"
  diff tests/fixtures/simple.expected.md "$TMP_MD" || true
fi
rm -f "$TMP_MD"

# --export-md - to stdout
STDOUT_OUT=$("${RUN[@]}" tests/fixtures/simple.csv --export-md - 2>/dev/null)
EXPECTED=$(cat tests/fixtures/simple.expected.md)
if [ "$STDOUT_OUT" = "$EXPECTED" ]; then
  pass "--export-md - writes Markdown to stdout"
else
  fail "--export-md - stdout differs from fixture"
fi

# --list-extensions runs without crashing (may print "no extensions installed")
if "${RUN[@]}" --list-extensions >/dev/null 2>&1; then
  pass "--list-extensions exits clean"
else
  fail "--list-extensions exited non-zero"
fi

# ── Summary ────────────────────────────────────────────────────────

echo
echo "$PASS passed, $FAIL failed"
[ "$FAIL" -eq 0 ]

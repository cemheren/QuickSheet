#!/usr/bin/env bash
# Scheduled runner for grow-quicksheet skill (hourly via cron)
# Runs Copilot CLI in non-interactive autopilot mode to execute one growth action.

set -euo pipefail

REPO_DIR="/home/akif/Projects/QuickSheet"
COPILOT="/home/akif/.local/bin/copilot"
LOG_DIR="$REPO_DIR/.agents/skills/grow-quicksheet"
RUN_LOG="$LOG_DIR/cron-runs.log"

cd "$REPO_DIR"

echo "$(date -Iseconds) — Starting grow-quicksheet run" >> "$RUN_LOG"

"$COPILOT" \
  -p "grow quicksheet" \
  --autopilot \
  --allowedTools "bash(read),bash(write),edit,create,view,grep,glob,web_search,web_fetch,github-mcp-server-get_file_contents,github-mcp-server-search_code" \
  2>&1 | tail -20 >> "$RUN_LOG"

echo "$(date -Iseconds) — Run complete" >> "$RUN_LOG"
echo "---" >> "$RUN_LOG"

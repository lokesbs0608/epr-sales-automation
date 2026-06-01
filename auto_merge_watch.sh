#!/usr/bin/env bash
# auto_merge_watch.sh
# Watches the orchestrator/worker processes and rebuilds FINAL_PP.xlsx /
# FINAL_PET.xlsx whenever:
#   - the run finishes (no more workers), OR
#   - it's interrupted (workers disappear), OR
#   - periodically while running (so you always have a fresh snapshot).
#
# Safe to leave running in a second terminal. Ctrl+C to stop watching.
# Usage: bash auto_merge_watch.sh [interval_seconds]

cd "$(dirname "$0")"
INTERVAL="${1:-60}"

running() { pgrep -f run_registered.js >/dev/null 2>&1 || pgrep -f 'node.*orchestrate' >/dev/null 2>&1; }

echo "[watch] started. Re-merging every ${INTERVAL}s while the run is active, plus a final merge when it ends."
WAS_RUNNING=0
while true; do
  if running; then
    WAS_RUNNING=1
    node merge_chunks.js >/tmp/auto_merge.log 2>&1
    echo "[watch $(date +%H:%M:%S)] snapshot merged (run active). $(grep -E 'unique rows|valid EPR' /tmp/auto_merge.log | tr '\n' ' ')"
    sleep "$INTERVAL"
  else
    # one final merge after the run ends/interrupts
    node merge_chunks.js >/tmp/auto_merge.log 2>&1
    echo "[watch $(date +%H:%M:%S)] FINAL merge done (run not active)."
    sed -n '/=== /,$p' /tmp/auto_merge.log
    if [ "$WAS_RUNNING" -eq 1 ]; then
      echo "[watch] run has ended — FINAL_PP.xlsx / FINAL_PET.xlsx are up to date. Stopping watcher."
      break
    fi
    echo "[watch] (no run active yet; merged current chunks). Waiting ${INTERVAL}s..."
    sleep "$INTERVAL"
  fi
done

#!/usr/bin/env bash
# Run once after: gh auth login
# Creates https://github.com/srbennett3/bis-planner and pushes main.
set -euo pipefail
cd "$(dirname "$0")/.."
unset CI
export GITHUB_TOKEN="${GITHUB_TOKEN:-}"
if ! command -v gh >/dev/null 2>&1; then
  echo "Install GitHub CLI: brew install gh"
  exit 1
fi
if ! gh auth status >/dev/null 2>&1; then
  echo "Not logged in. Run: gh auth login"
  exit 1
fi
REPO="srbennett3/bis-planner"
if git remote get-url origin >/dev/null 2>&1; then
  echo "Remote origin already set. Pushing..."
  git push -u origin main
else
  gh repo create "$REPO" --public --source=. --remote=origin --push \
    --description "TBC Classic BIS gear planner (Python generator + Google Apps Script)"
fi
echo "Done: https://github.com/$REPO"

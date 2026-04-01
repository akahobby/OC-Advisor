#!/usr/bin/env bash
set -euo pipefail

repo_root="$(git rev-parse --show-toplevel)"
hooks_dir="$repo_root/.githooks"

if [ ! -d "$hooks_dir" ]; then
  echo "Missing $hooks_dir" >&2
  exit 1
fi

chmod +x "$hooks_dir/commit-msg"
git config core.hooksPath .githooks

echo "Installed local hooks from .githooks"
echo "Current hooksPath: $(git config --get core.hooksPath)"

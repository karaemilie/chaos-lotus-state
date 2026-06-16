name: Process drain.json into beast

# Fires when drain.json is updated by the worker (after every tap on the wheel),
# OR on manual dispatch from the Actions tab for debugging.
on:
  push:
    branches: [main]
    paths:
      - 'drain.json'
  workflow_dispatch:

# Grant write permission to the workflow's default GITHUB_TOKEN
# (Defaults to read-only since 2023 — must opt in.)
permissions:
  contents: write

# Concurrency: if a new run starts while one is in progress, queue it.
# Don't cancel — we want to process every batch (the new state of drain.json
# will be picked up by the queued run).
concurrency:
  group: process-drain
  cancel-in-progress: false

jobs:
  process:
    runs-on: ubuntu-latest
    steps:
      - name: Checkout chaos-lotus-state (self)
        uses: actions/checkout@v4
        with:
          # Default GITHUB_TOKEN — has push access to THIS repo only
          token: ${{ secrets.GITHUB_TOKEN }}

      - name: Setup Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Install dependencies
        run: pip install openpyxl

      - name: Process drain
        env:
          CROSS_REPO_TOKEN: ${{ secrets.CROSS_REPO_TOKEN }}
        run: python .github/actions-scripts/process_drain.py

      - name: Commit processed files (drain + receipt) with rebase-retry
        run: |
          # The worker (outside this Action) commits drain.json mirrors between
          # our checkout and our push. A plain `git push` then fails with a
          # non-fast-forward rejection, and the processed/cleared drain never
          # lands — so completions bounce back on the next mirror. Fix: stage
          # our changes, then pull --rebase (re-applying our commit on top of
          # whatever the worker pushed) and push, retrying a few times to win
          # any race. last_sync.json is also committed here so the public
          # verification receipt actually persists.
          git config user.name "github-actions[bot]"
          git config user.email "github-actions[bot]@users.noreply.github.com"

          # Stage whatever the processor changed (drain always; receipt if written).
          git add drain.json last_sync.json 2>/dev/null || git add drain.json

          if git diff --cached --quiet; then
            echo "No staged changes — nothing to commit."
            exit 0
          fi

          git commit -m "🤖 drain.json processed + receipt updated"

          # Rebase-and-push with retries to survive concurrent worker mirrors.
          # Under rapid taps, many runs fire near-simultaneously; losing the
          # push race is EXPECTED and self-healing (our processed drain is safe
          # locally, and the next worker mirror triggers a fresh run that
          # reprocesses anything outstanding). So a lost race must NOT fail the
          # run (exit 1 spams "Run failed" emails for a non-problem). We retry
          # generously with jitter, then exit 0 gracefully if still racing.
          n=0
          until [ "$n" -ge 8 ]; do
            # Bring in any commits the worker pushed since our checkout, replaying
            # our commit on top. --autostash guards any stray unstaged changes.
            if git pull --rebase --autostash origin main; then
              if git push origin main; then
                echo "✅ Pushed processed drain on attempt $((n+1))"
                exit 0
              fi
            fi
            n=$((n+1))
            # Jittered backoff (2-5s) so concurrent runs don't lock-step collide.
            jitter=$((RANDOM % 4 + 2))
            echo "⚠️  push race (attempt $n/8) — retrying in ${jitter}s..."
            sleep "$jitter"
          done

          # Still racing after 8 tries: the drain we processed is preserved and
          # will be picked up by a subsequent run. Exit SUCCESS so GitHub doesn't
          # email a failure for this benign, self-healing condition.
          echo "ℹ️  Could not win push race after 8 attempts — leaving for the"
          echo "   next run to reprocess (drain is preserved). Exiting cleanly."
          exit 0

# SACCR working instructions

## Active setup milestone

- Until milestone `v0.0.1` is closed, all development and setup changes belong
  on `release/0.0.1`, as directed by the owner on 2026-10-06.
- Do not commit to or merge into `main` unless the owner explicitly requests
  that specific main-branch change. Closing the milestone does not itself
  authorize merging, tagging or publishing a release.
- If a task branch is needed, create it from `release/0.0.1` and target that
  branch with its pull request. Never assume GitHub's default PR base is right.
- Preserve both branch histories when reconciling existing setup work. Do not
  force-push, reset shared history or delete the active release branch.
- Every issue must be assigned to `danielep71`, have at least one of `P1`,
  `P2`, `P3`, and have the relevant milestone. Prefer one clear priority.
- Keep SACCR private. Keep Dependabot low-impact alert auto-dismissal disabled.
- Run `python tools/check.py --ci` on a clean committed candidate; for a PR
  use `--base <base-sha>`. Verify hosted checks and preserve evidence before
  marking work complete. Report static checks and actual Excel runs separately.
- Follow [repository governance](docs/GOVERNANCE.md). Its review process is a
  working convention, not a claim that GitHub enforces branch protection.

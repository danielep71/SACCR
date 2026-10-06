# Repository governance

Owner and maintainer: **@danielep71**. Setup milestone: **v0.0.1**.
This policy implements the owner's working-branch decision of 2026-10-06 and
the scope of [issue #2](https://github.com/danielep71/SACCR/issues/2).

## Branch policy during setup

| Branch | Purpose and allowed work |
| --- | --- |
| `release/0.0.1` | Active integration branch for all development and setup until milestone v0.0.1 closes. |
| `main` | Default branch. Receives `release/0.0.1` only through an owner-requested PR merged with a merge commit; last integrated by PR #14 (`ff71489`). Nothing is committed to main directly. |
| Short-lived task branches | Every change: branch from release/0.0.1, open the PR against release/0.0.1 and squash-merge it. Use a descriptive name such as `fix/<issue>-<slug>`, `docs/<slug>` or `chore/<slug>`. |
| Other `release/*` branches | Reserved for explicitly agreed future milestones; do not infer a new development target from a version number. |

The default branch remains `main`; this does not make it the active development
branch. Check the PR base explicitly. Closing a milestone does not publish a
tag, release or workbook, or authorize an automatic merge into main. The owner
must agree the final integration and subsequent working branch.

Task PRs into the release branch are squash-merged. Integration of the release
branch into main uses a PR merged with a merge commit once hosted checks are
green, so provenance remains visible; afterwards the release branch is
fast-forwarded to main, so both point to the merge commit. Never force-push or
reset shared branches. Merged task branches are deleted automatically; preserve
the active release branch, and if a release-to-main merge deletes it, recreate
it from main at the merge commit. Do not open a release-to-main PR just to
record progress during this setup phase.

## Ownership, authorization and review

The owner sets scope, priorities, exceptions and release decisions. Contributors
work within the authorized issue. Every change reaches the active release
branch through a PR; nothing is committed directly to the release branch or
main. Inspect the diff and run the applicable checks before pushing. An
author's self-check is not an independent review.

No change to visibility, paid subscriptions, security-sensitive access or
main-branch scope is implied by an ordinary development task. Raise a concrete
exception when needed rather than bypassing a requirement. Keep unresolved
review findings and failed checks visible; do not call unavailable validation
successful.

Minimum review checklist, before integration or recording completion:

- [ ] The issue is assigned to danielep71, has a priority and the right milestone.
- [ ] Scope and acceptance criteria match the diff; the destination is the active release branch.
- [ ] Interfaces, units, error behavior, cleanup and compatibility impact are explained where relevant.
- [ ] Applicable local tests pass and the exact committed candidate is identified.
- [ ] Applicable hosted checks pass; PR validation covers the proposed range and reports are retained.
- [ ] Review findings are resolved or explicitly accepted with a reason and owner decision where needed.
- [ ] Documentation and Unreleased entries reflect actual behavior and validation.
- [ ] No credentials, client data, unrelated binary outputs or undeclared dependencies are introduced.
- [ ] Static results are separated from real Excel compilation/execution evidence.

Local validation (`python tools/check.py`) precedes every push; run
`python tools/check.py --ci` on the clean committed candidate, with
`--base <base-sha>` for a PR. Hosted checks must be green before a PR is
merged. Report static checks and real Excel runs separately. If a hosted run fails after a
merge, record the failure and fix or revert it through a new PR before closing
the issue. Never rewrite shared history to conceal the failed candidate.

## Issue metadata and completion

Every issue must have assignee **danielep71**, at least one priority label
**P1/P2/P3**, and a reference milestone; use one clear priority unless a
documented reason requires otherwise. Current setup work belongs to v0.0.1.
Include objective, acceptance checklist, dependencies and required evidence.

### What is automatic and what is not

| Field | How it is set |
| --- | --- |
| Assignee `danielep71` | Set by every issue form |
| Type label (`bug`, `enhancement`, `documentation`) | Set by the form |
| Priority | Each form sets a default (`P2`, or `P3` for documentation); triage confirms or changes it |
| Milestone | **Manual.** GitHub issue forms cannot set a milestone |

Forms apply only to issues opened in the web interface; blank issues are
disabled there. Issues created through the API, the CLI or another tool bypass
the forms and get no metadata automatically. Nothing blocks an issue that lacks
metadata: this is a **manual rule backed by the check below**, not enforcement.

### Triage check

Run these three searches when an issue is opened and before closing a
milestone. Each must return no results.

- [Open issues without a milestone](https://github.com/danielep71/SACCR/issues?q=is%3Aissue%20is%3Aopen%20no%3Amilestone)
- [Open issues without an assignee](https://github.com/danielep71/SACCR/issues?q=is%3Aissue%20is%3Aopen%20no%3Aassignee)
- [Open issues without a priority label](https://github.com/danielep71/SACCR/issues?q=is%3Aissue%20is%3Aopen%20-label%3AP1%20-label%3AP2%20-label%3AP3)

The issue forms and the pull-request template are read by GitHub from the
default branch, so a change to them takes effect after the next integration
into `main`.

### Closing an issue

Close an issue only when every acceptance criterion has linked evidence: the
merged PR, the green check run and, for VBA, the Excel result. Record that
evidence in a closing comment. A blocked Excel run keeps the issue open.
Milestone closeout is tracked in
[issue #13](https://github.com/danielep71/SACCR/issues/13).

## Current GitHub controls and limitations

Observed for this private personal repository on 2026-10-06. During issue #2,
the owner upgraded the account: GitHub's Licensing page now confirms
**Current GitHub base plan: GitHub Pro**. The earlier Free-plan restriction is
historical, not a current blocker. Availability and activation are distinct.

| Control | Observed state and working response |
| --- | --- |
| Repository visibility | Private. Preserve it; public visibility is not a workaround for missing features. |
| Branch rulesets | Configured by the owner on 2026-10-06. The repository's Settings → Rules page is authoritative for which branches and rules apply; keep following the review checklist for what rulesets cannot check. |
| Automatic head-branch deletion | Enabled by the owner on 2026-10-06. Merged PR branches are deleted automatically. |
| Wiki | Pro enables Wiki for private personal repositories, but SACCR's Wiki remains disabled in this baseline. Keep authoritative documentation in versioned docs/. |
| Advanced code/secret scanning | Not offered in the current repository security settings. Do not claim CodeQL, secret scanning or push protection are active. Review code and credentials manually; this is not equivalent automated coverage. |
| Dependency graph and Dependabot alerts | Enabled. Review alerts rather than assuming a clean dependency graph proves application security. |
| Low-impact development-dependency alert auto-dismissal | Disabled by explicit owner decision. Keep those alerts visible. |
| Automatic Dependabot updates | Weekly update PRs for GitHub Actions, aimed at the active release branch, from `.github/dependabot.yml` once it is on `main`. Every update is reviewed and merged manually; see [tools/README.md](../tools/README.md#dependency-updates). Alerts and update PRs are separate features. |
| Repository integrity CI | Runs on main/release pushes and PRs, read-only token and full-SHA action pins. Also lints the Python tooling (Ruff, strict mypy) and the workflows (actionlint). Reports are retained for 30 days. Static checks only; no Excel execution. |
| Labels | Twenty-label catalogue with sync and read-only drift workflows. See LABELS.md for permissions and triggers. |

GitHub Pro enables capabilities such as protected branches, required reviewers
and Wiki in private personal repositories. It does not automatically configure
those controls and must not be represented as enabling all advanced security
products. The private repository's security settings still do not offer the
advanced code/secret-scanning sections after the upgrade.
Reference: [GitHub plans](https://docs.github.com/en/get-started/learning-about-github/githubs-plans).
Recheck feature availability before future configuration changes. Record the
ruleset configuration and the remaining Wiki decision under the final setup
review in issue #13. Issue #2 itself changes
documentation and the release-branch baseline, not account billing or visibility.

## Label workflows

Main and the release branch carry the same label catalogue, scripts and
workflows. PRs affecting them run offline validation. The automatic live sync
is triggered by changes to its inputs on main; scheduled drift uses the default
branch. Changes merged into the release branch alone do not trigger live sync.

If an authorized label-policy change is developed on release, validate it
locally and via PR checks, review removals and their issue associations, then
use an explicitly requested manual sync on that reviewed release ref when
live reconciliation is needed. Do not modify main solely to trigger sync.
The label workflows do not automatically assign priorities or milestones.

## Reconciled setup baseline

Before this reconciliation:

- main: `28af65c007b55434c1f95b7035e72b03ecdffe8e`;
- release/0.0.1: `2f8ef9d89702185c2afc1762157b71c07fe70d02`.

Both contained the same tooling and static workflow. Release lacked seven
label-policy, script, workflow and documentation files present in main.
The governance integration preserves both commits as parents, first the
release tip and then main, and imports those seven files unchanged into
release. Main stayed at the recorded tip at that point. The changelog states
release is the active setup branch, and the README links this policy.

PR #14 later merged the release branch into main with merge commit
`ff714896eabfb32a91659b7d80aa9ad9b5607496`. Immediately after that merge both
branches pointed to that commit, and hosted checks passed on it for both; the
release branch has moved on since.

The accepted integration commit, hosted run, branch comparison and settings
observations are recorded in issue #2. Retain necessary closeout evidence
before the workflow artifacts expire; final reproducibility and release
certification remain the responsibility of issue #13.

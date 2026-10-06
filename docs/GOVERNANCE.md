# Repository governance

Owner and maintainer: **@danielep71**. Setup milestone: **v0.0.1**.
This policy implements the owner's working-branch decision of 2026-10-06 and
the scope of [issue #2](https://github.com/danielep71/SACCR/issues/2).

## Branch policy during setup

| Branch | Purpose and allowed work |
| --- | --- |
| `release/0.0.1` | Active integration branch for all development and setup until milestone v0.0.1 closes. |
| `main` | Preserved default-branch baseline. A specific owner instruction is required for any commit or merge into main during setup. |
| Short-lived task branches | When needed, branch from release/0.0.1 and target release/0.0.1 in the PR. Use an issue-related name such as chore/2-governance. |
| Other `release/*` branches | Reserved for explicitly agreed future milestones; do not infer a new development target from a version number. |

The default branch remains `main`; this does not make it the active development
branch. Check the PR base explicitly. Closing a milestone does not publish a
tag, release or workbook, or authorize an automatic merge into main. The owner
must agree the final integration and subsequent working branch.

Use a non-rewriting merge for integration so provenance remains visible. Never
force-push or reset shared branches. Delete only disposable task branches after
their work is integrated; preserve the active release branch. Do not open a
release-to-main PR just to record progress during this setup phase.

## Ownership, authorization and review

The owner sets scope, priorities, exceptions and release decisions. Contributors
and agents work within the authorized issue. Routine scoped changes may be
committed directly to the active release branch when covered by the owner's
task instruction; use a PR for changes requiring review, collaboration or
validation of the PR path. In both cases, inspect the diff and run the same
applicable checks. An agent's self-check is not an independent human review.

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

For an authorized direct release commit, local validation precedes the push;
hosted validation follows it. If the hosted run fails, record the failure and
fix or revert with a new commit before closing the issue. Never rewrite shared
history to conceal the failed candidate.

## Issue metadata and completion

Every issue must have assignee **danielep71**, at least one priority label
**P1/P2/P3**, and a reference milestone; use one clear priority unless a
documented reason requires otherwise. Current setup work belongs to v0.0.1.
Include objective, acceptance checklist, dependencies and required evidence.

These are mandatory working rules, not currently an automated metadata gate.
Check them at creation and review. Template and enforcement improvements are
tracked in [issue #11](https://github.com/danielep71/SACCR/issues/11).
Close an issue only when its checklist and evidence agree. A blocked Excel
run remains pending. Milestone closeout is tracked separately in
[issue #13](https://github.com/danielep71/SACCR/issues/13).

## Current GitHub controls and limitations

Observed for this private personal repository on 2026-10-06. During issue #2,
the owner upgraded the account: GitHub's Licensing page now confirms
**Current GitHub base plan: GitHub Pro**. The earlier Free-plan restriction is
historical, not a current blocker. Availability and activation are distinct.

| Control | Observed state and working response |
| --- | --- |
| Repository visibility | Private. Preserve it; public visibility is not a workaround for missing features. |
| Branch protection/rulesets | Pro now makes protection available; the active branches are still unprotected and no classic rule is configured. Follow the review checklist manually until controls are configured and verified. A green check does not itself prohibit a merge or direct push. |
| Wiki | Pro enables Wiki for private personal repositories, but SACCR's Wiki remains disabled in this baseline. Keep authoritative documentation in versioned docs/. |
| Advanced code/secret scanning | Not offered in the current repository security settings. Do not claim CodeQL, secret scanning or push protection are active. Review code and credentials manually; this is not equivalent automated coverage. |
| Dependency graph and Dependabot alerts | Enabled. Review alerts rather than assuming a clean dependency graph proves application security. |
| Low-impact development-dependency alert auto-dismissal | Disabled by explicit owner decision. Keep those alerts visible. |
| Automatic Dependabot updates | Not configured in this baseline; Actions update configuration is tracked by issue #6. Alerts and automated update PRs are separate features. |
| Repository integrity CI | Runs on main/release pushes and PRs, read-only token and full-SHA action pins. Reports are retained for 30 days. Static checks only; no Excel execution. |
| Labels | Twenty-label catalogue with sync and read-only drift workflows. See LABELS.md for permissions and triggers. |

GitHub Pro enables capabilities such as protected branches, required reviewers
and Wiki in private personal repositories. It does not automatically configure
those controls and must not be represented as enabling all advanced security
products. The private repository's security settings still do not offer the
advanced code/secret-scanning sections after the upgrade.
Reference: [GitHub plans](https://docs.github.com/en/get-started/learning-about-github/githubs-plans).
Recheck feature availability before future configuration changes. Record the
remaining protection/Wiki configuration under the final setup review in issue
#13; do not mark it enabled just because Pro is active. Issue #2 itself changes
documentation and the release-branch baseline, not account billing or visibility.

## Label workflows while main is preserved

The release branch carries the same label catalogue, scripts and workflows as
main. PRs affecting them run offline validation. The existing automatic live
sync is triggered by changes on main; scheduled drift uses the default branch.
Copying the files to release does not change those trigger semantics.

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
release. Main remains at the recorded tip. The changelog states release is
the active setup branch; the README links this policy and AGENTS.md records
the owner's instructions for future agents.

The accepted integration commit, hosted run, branch comparison and settings
observations are recorded in issue #2. Retain necessary closeout evidence
before the workflow artifacts expire; final reproducibility and release
certification remain the responsibility of issue #13.

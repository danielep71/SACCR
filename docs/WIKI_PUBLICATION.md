# Wiki Source and Publication

The GitHub Wiki is a **derived publication surface**. Every maintained Wiki page
has reviewed source under `docs/wiki/`; an edit made only in the GitHub Wiki UI
is drift and is not authoritative.

## Source contract

`docs/wiki/catalogue.json` defines the ordered page set and each page's
repository authority. `Home.md` is first and `_Sidebar.md` is generated from
the catalogue. Each page contains one **Guide, not policy** notice linking to its
registered authority.

Run:

```sh
python tools/check_wiki.py --root .
python tools/test_wiki.py -v
```

After changing the catalogue, refresh the sidebar with:

```sh
python tools/check_wiki.py --root . --write-sidebar
git add docs/wiki/_Sidebar.md
python tools/check_wiki.py --root .
```

The normal `python tools/check.py` gate also runs the Wiki tests and source check.

## Publication

Publication requires a clean, committed source tree. The exported Wiki pins
repository-document links to the exact source SHA and writes
`Wiki-Source.json` with the repository, source SHA and SHA-256 of every page.

On Windows, double-click:

```text
tools\Publish-SACCR-Wiki.cmd
```

The publisher derives the Wiki URL from `origin`, clones or fast-forwards a
sibling Wiki checkout, refuses to overwrite unmanaged files, replaces the whole
managed page set, removes obsolete managed pages, commits and pushes only when
bytes changed, and then verifies a fresh clone byte-for-byte.

The default Wiki checkout is a sibling named `<source-folder>.wiki`. Direct
PowerShell use may override it with `-WikiPath`.

## Drift

`.github/workflows/wiki-drift.yml` periodically clones the public Wiki, reads
its recorded source commit, checks out that exact commit and performs the same
byte comparison. A missing manifest, online-only edit or mismatched page is
reported as drift.

If an existing Wiki page is not yet represented in `docs/wiki/`, migrate it
into the catalogue first. The publisher stops rather than silently deleting an
unmanaged page.

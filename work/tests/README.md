# tests

Offline. Run from `work/`:

    python3 -m tests                 all
    python3 -m tests.test_gdeck      one file

| File | Covers |
|---|---|
| `test_gdeck.py` | ownership rules on a saved real deck |
| `test_gwrite.py` | guard, rewrites, revision-clash retry |
| `test_gsteps.py` | news selection, TOC, preflight, skeleton, on a lived-in deck |
| `test_ids.py` | object id scheme |
| `test_layout.py` | layout, benchmark measurements, style constants |
| `test_deck.py` | inline markup, topic matching, parser |
| `runner.py` | runs the `test_*` functions of one module |
| `fixtures/` | real `presentations.get` responses the tests read |

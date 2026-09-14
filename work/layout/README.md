# layout

Pure content and geometry. No network, no Google.

| Module | Owns |
|---|---|
| `deck_parser.py` | `Block`, `Section`, inline markup (`split_markup`, `strip_markup`, `is_link`), `same_topic`. Also the Markdown deck grammar of the earlier workflow |
| `deck_layout.py` | pagination, font ladder, every rectangle (`paginate`, `Rect`, constants) |
| `text_metrics.py` | how wide text renders in Calibri (`text_width`) |
| `bench_page.py` | benchmark page sizes, vendor colours, cutoff note |

Depends on nothing else in the project.

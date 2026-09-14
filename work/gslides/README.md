# gslides

Everything that touches a live Google Slides deck.

| Module | Owns | Public API |
|---|---|---|
| `client.py` | sign-in, settings, project paths | `services()`, `settings()`, `log()`, `ROOT`, `CONFIG_DIR`, `DATA_DIR` |
| `deck.py` | read model and ownership rules | `read_deck()`, `parse_deck()`, `Deck.editable()`, `find_decks()`, `seminar_date()` |
| `write.py` | safe writes | `run_plan()`, `refresh_text()`, `refresh_topic()`, `refresh_picture()`, `replace_line()`, `replace_owned()`, `guard()` |
| `render.py` | content to Slides requests | `render_content()`, `render_bench()`, `render_toc_columns()`, `Body`, `make_box()` |
| `api.py` | single request dictionaries | `textbox()`, `insert_text()`, `style_span()`, `picture()`, ... |
| `ids.py` | object ids, fixed slide names, ownership prefix | `block_key()`, `shape_id()`, `owned()`, `PAGE_*`, `TOPIC_*` |
| `host.py` | short-lived public picture URLs on Drive | `Host.put()`, `Host.clean()` |
| `step.py` | what every g*_ command does first | `start()`, `with_pictures()`, `page_or_stop()` |

Depends on `layout` and `sources`. Only `client`, `read_deck`,
`write.apply` and `host` call Google; everything else is
pure and tested offline.

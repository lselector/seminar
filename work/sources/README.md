# sources

Outside content, gathered on this machine.

| Module | Owns |
|---|---|
| `leaderboard.py` | LM Arena top 25 into `data/leaderboard.json` and two xlsx. Run: `python3 -m sources.leaderboard` |
| `pictures.py` | one picture: download or screenshot, then clean (`prepare`, `challenged`) |
| `fetch_images.py` | download and headless Chrome helpers used by `pictures.py` |
| `clean_images.py` | the ImageMagick conversion used by `pictures.py` |

`fetch_images.py` and `clean_images.py` still have command
lines from the earlier Markdown workflow; nothing calls them.

Depends on `layout` only.

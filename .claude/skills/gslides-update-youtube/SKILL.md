---
name: gslides-update-youtube
description: Refresh the subscriber and video counts and the channel screenshot on the "Weekly videos every Friday" slide of the live Google Slides seminar deck, changing only the counts line so the author's wording stays. Use when asked to update the YouTube, channel or subscribers slide in the Google Slides deck.
---

# Update the YouTube slide

Run from `work/`. The rules are in `work/README.md`.

```bash
python3 g6_update_youtube.py          # or: ... 2026-09-25
```

It reads the counts from https://www.youtube.com/@lev-selector
and replaces only the line containing `subscribers`, then
retakes the channel screenshot.

If the log says `counts not found`, open the channel page
(WebFetch), read the two numbers, and pass them:

```bash
echo '{"subscribers": "7.51K", "videos": "337"}' > /tmp/yt.json
python3 g6_update_youtube.py --json /tmp/yt.json
```

Report the counts now shown. If the box is filled, the
author froze it; say so.

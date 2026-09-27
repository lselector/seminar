---
name: gslides-update-youtube
description: Refresh the subscriber and video counts, the channel screenshot, and the screenshots of the channel's Videos and Shorts tabs on the "Weekly videos every Friday" slide of the live Google Slides seminar deck, changing only the counts line so the author's wording stays. Use when asked to update the YouTube, channel, videos, shorts or subscribers slide in the Google Slides deck.
---

# Update the YouTube slide

Run from `work/`. The rules are in `work/README.md`.
Without a date the step works on this week's deck: on a
Friday that is today's until 3 pm US Eastern (the seminar
is over), then next week's. Say which deck the log names.
If the user names a date, pass it to every command.

```bash
python3 g6_update_youtube.py          # or: ... 2026-09-25
```

It reads the counts from https://www.youtube.com/@lev-selector
and replaces only the line containing `subscribers`. Then it
retakes the pictures:

| Picture | Shot |
|---|---|
| alt text `auto: youtube-videos` | channel header and the Videos tab (`/videos`) |
| alt text `auto: youtube-shorts` | the Shorts tab (`/shorts`), from the tabs down |
| the script's own channel picture | the channel page |

Each tab is shot in its frame's width:height shape, so it
fills the frame where it stands.

Page 5 is the author's own design: a promo box and the two
tab screenshots. `g1_new_deck.py` gives every new deck a
copy of last week's page 5, marks included, so its counts
and screenshots are last week's until this step runs.
Pictures without a mark are never replaced.

That line is kept current by request, so it is rewritten in
whichever box on the slide holds it, filled or not, and in a
promo box the author wrote themselves. The rest of the box
stays as they left it. If the promo box is gone altogether,
the step draws it again from `config/skeleton.json` with the
current counts, and says so.

If the counts can't be read from the channel page, the log
says `counts not found`. Open the channel page (WebFetch),
read the two numbers, and pass them:

```bash
echo '{"subscribers": "7.51K", "videos": "337"}' > /tmp/yt.json
python3 g6_update_youtube.py --json /tmp/yt.json
```

| Log line | Meaning |
|---|---|
| `no picture has the alt text 'auto: youtube-...'` | Tell the author to add that mark to the picture (right-click, Alt text, Description) |
| `picture ...: not available, the old picture stays` | The screenshot failed; the old one is kept |

Report the counts now shown and which screenshots were
retaken. Say so if the log reports that the box was filled,
or that the promo box had to be drawn again.

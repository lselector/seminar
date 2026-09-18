---
name: gslides-update-youtube
description: Refresh the subscriber and video counts and the channel screenshot on the "Weekly videos every Friday" slide of the live Google Slides seminar deck, changing only the counts line so the author's wording stays. Use when asked to update the YouTube, channel or subscribers slide in the Google Slides deck.
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
and replaces only the line containing `subscribers`, then
retakes the channel screenshot.

Page 5 is the author's own design: a promo box and two
channel screenshots. `g1_new_deck.py` gives every new deck
a copy of last week's page 5, so its counts are last
week's until this step runs. The screenshots are the
author's and are never replaced.

That line is kept current by request, so it is rewritten in
whichever box on the slide holds it, filled or not, and in a
promo box the author wrote themselves. The rest of the box
stays as they left it. If the promo box is gone altogether,
the step draws it again from `config/skeleton.json` with the
current counts, and says so.

If the log says `counts not found`, open the channel page
(WebFetch), read the two numbers, and pass them:

```bash
echo '{"subscribers": "7.51K", "videos": "337"}' > /tmp/yt.json
python3 g6_update_youtube.py --json /tmp/yt.json
```

Report the counts now shown. Say so if the log reports that
the box was filled, or that the promo box had to be drawn
again.

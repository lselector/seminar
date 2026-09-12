
## Lev Selector - Weekly Seminar

This GitHub repository is used to 
create and store slides for weekly seminars.

Slides are saved in "pptx" (or "pdf") format 
in subdirectories by year:

- 2021_1_data_science
- 2021_2_data_architect
- 2022
- 2023
- 2024
- 2025
- 2026

The files are tagged by date in format YYYY-MM-DD,
for example: 2026/2026-09-11-AI-Updates.pptx

The videos of the seminar are posted on youtube.
Search for Lev Selector
 - https://www.youtube.com/@lev-selector

Subdirectory "work" is for generating
those slides using AI (Claude Code)

One Markdown file per week holds all the
content, and Python scripts fetch the images
and build the PPTX from it. Slide layout is
fully automatic - the Markdown file holds no
positions, sizes, or slide breaks.

Two documents cover it:

- work/README.md - how to use it, step by step
- work/ADD.md - the design and the reasoning

The short version. To build a week's deck in
Claude Code, run the skill:

```
/slides-update
```

To rebuild the PPTX after editing the md file
by hand, for example minutes before the
seminar:

```
cd work
python3 s3_make_pptx.py
```

When editing the md file by hand, mark your
own work so it is never rewritten:

- `<!-- locked -->` - I wrote this, do not change it
- `<!-- deleted -->` - I dropped this topic, do not bring it back

See work/README.md for the rest, including how
to add your own images.

Two files (seminar_pdf.txt, seminar_ppt.txt)
are generated automatically by extracting
text strings from all files in this project.
These two files are convenient for searching
using the unix "grep" command.

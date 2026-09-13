
-----------------------------

This directory is a project to create slides 
for weekly news and updates about progress 
related to AI (Artificial Intelligence).

Please review the two recent files:

2026-09-11-AI-Updates.pptx - the latest seminar
2026-09-18-AI-Updates.pptx - the next future seminar

The numbers in file names are dates in format YYYY-MM-DD

I want to establish the following process:
- make a list of news in a text (md) file
  (the list should include short 80-100words descriptions)
- download relevant images into images_raw directory
- clean/resize the images using python script

- create a Markdown file following the structure of the 
  examples pptx files in this directory.
  It shoudl be landscape orientation (for slides and PPTX)

  It should consist of:
    Table of contents
    Benchmarks page
    Intelligence index page
    several pages of the most important news
    a page about my youtube channel @lev-selector
    more pages about news
    a page about layoffs and hiring statistics
    page about the author (me) - and thank you page

Does this plan makes sense?
Please create a concise ADD.md (Architecture Design Document)
for this process

-----------------------------

Please do the following changes:

- in ADD - allow automatic page layout. Human should not do this. Each piece of news should consists of text and an illustrative image. 

- create python script to generate PPTX from MD

- change s2_clean_images.py as needed

use leaderboard.py script to download recent benchmarks

- in ADD - account for handling pages for Benchmarks, Intelligence Index, and Jobs

-----------------------------

Good first version. Thank you.

Now, for future work please create a skill
"slides_update" which will create or update the slide deck by working on a md file, and then converting it to pptx file.

The skill may receive the date for the slides.
If no date is provided, it should calculate the date as next Friday
If tody is Friday - it should work on today's slides (unless another date is explicitly given)

If the slides files don't exist - the skill should create them with names:

    YYYY-MM-DD-AI-Updates.md
    YYYY-MM-DD-AI-Updates-generated.pptx

If the files already exist - the skill should update them as needed.

A human author should be able to make changes to the md file in a way so that AI will not revert them

-----------------------------

Just before the seminar a human author should be able to make last minute manual changes in the "md" file - and then run a short python script to regenerate PPTX. Will the s3_make_pptx.py script do the right thing?

-----------------------------

A human author can make 3 main types of edits in "md" file: add, update, delete

If human had updated/deleted something - there shoudl be a way to indicate that the particular topic was updated or deleted by a human - and should not be changed

-----------------------------

Please add the README.md file in this "work" directory with instructions for human
- how to run the skill
- how to do manual edits
- how to add image manually
- how to do final touches - and regenerate the PPTX
- etc.

-----------------------------

I want to change all file names from "AI-Updates" to "AI-News"

-----------------------------

Want to make a change to the workflow for deleting content.
Suppose I am manually editing file 2026-09-18-AI-News.md
and want to remove a particular item.
Instead of marking it and living in the same file,
I want to put it into a separate file named 
2026-09-18-AI-News-deleted.md

So the skill should recognize that 
this topic was removed by a human
and should not be recreated in the main file

-----------------------------

I want to be able to delete the itmes using two methods.

First method - delete by hand from md file - and paste it into the "-deleted" file

Second method - make it as deleted in the md file, and then run special python script which will move all deleted items into the "-deleted" file

-----------------------------

I used to create manually in Google Slides.
The downloading them as PPTX file.
Then I was opening the PPTX, selecting one of the images - and selecting FILE >Compress Pictures...
Then selecting Picture Quality as email (96 ppi), and checkboxes to removed cropped areas and apply this to all images in the file.

This operation was routinely reducing the file size from 20..30 MBytes down to approx 2 MBytes.

How can we make sure that our new automation process (downloading and processing files, geneerating md and pptx files) will keep the final file size small? Please make necessary changes in the process.

-----------------------------

The two tables on page 2 of generated pptx (benchmarks) are too big and don't fit.
maybe we should convert them into plain text?
Maybe change the order of columns to this:

   colored marker - score - name

-----------------------------

Please recize all textboxes in the PPTX to remove extra spaces. 

Also fill them with light yellow-ish background color like it was in previous week PPTs.

Also please add red borders to all text boxes

-----------------------------

Please do two changes on the first page of the PPTX:

1. please remove these two lines:

Lev Selector, Ph.D.
Questions and discussion

2. Please change font size to 14

-----------------------------

Yes, **"Subscribe to the channel** should not be in the 1st (TOC) page

-----------------------------

I usually generate a relevant epigraph for the whole slide deck - and put it in red phone on the top right of the first TOC page. See how it is done in 2026-09-11-AI-News.pptx

Please add this functionality to our process:

Automatically generate a good short relevant original energetic and smart epigraph - and add it to the md document and to PPTX

-----------------------------

On the 1st page please remove empty spaces at the bottom of text boxes

Similarly, on the 2nd page (Benchmarks) please remove extra space on the right of the text boxes.

When filling the text boxes, please make text as bulleted list consisting of short concize factual statements

Also, please add red border to all images

Also make changes to the author page in the end: increase font size and add my portrait (Lev Selector)

-----------------------------

On 2nd page (Benchmarks) add the date when these benchmarks were published on the website

-----------------------------

On the "About the Speaker" page change the layout to follow how it was done in previous PPTs. Larger image, different positionning on the page.

One the last page (Thank You!)
you currently have only link to youtube channel.
Please also add the following two links for slides

Slides on GDrive: https://drive.google.com/drive/folder...

Slides on GitHub - https://github.com/lselector/seminar
    (click on pptx file, then on "raw" or download button on the right)

-----------------------------

The very last page - Thank you! page:

Change the font size to 40 
for the "Thank Your!" title text box 
and move this text box to the
center horizontally and then one third of the page down

Put the text box with links also at the center horizontally under the title, and remove te border and the 
filling color.

-----------------------------

On the last page plese remove the "Questions and discussion" subtitle - it is not needed.

-----------------------------

Please also make rearrangements in the "About the Speaker" page.

See the attached photo.

Put photo and text more in center both vertically and horizontally. Remove text box line and fill.

-----------------------------

On regular slides we have a title
on the top left. The text box for it currently takes all the width of the slide. Please adjust its width - remove the unnecessary width on the right. Do it on all regular slides.

-----------------------------

Currently regular slides contain text in text boxes with yellow background.

The font should be Calibri 12
The title of each box should be bold and red

If the textbox contains code,
it should be fixed width font with size 9 and color blue

-----------------------------

Please make adjustments to page 2 "Benchmarks"
The font size for text boxes on
this page should be smaller than usual - 9 pts

Also currently the two tables are snugged to the sides. Please put them closer to eahc other in the center of the page horizontally.

Also the text giving the date ("Votes counted through Sept 11, 2026")
should be in a regular yellow textbox with regular font size 12 and should not be hidden by other textboxes.

-----------------------------

On most pages we have two or one yellow text boxes on the left. Please distribute them vertically so that there will be some white space aronud them and from the top title

-----------------------------

Please change page "Weekly Videos Every Friday" to look more like the attached file. Make font bigger (22 text, 18 link)


-----------------------------



-----------------------------

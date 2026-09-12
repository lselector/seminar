---
name: humanize
description: Write prose that reads like a human wrote it, and rewrite AI-generated prose so it does. Covers English and Russian. Use when drafting any prose (book chapters, forewords, blurbs, captions, articles, documentation narrative) to avoid AI tells from the first draft. Also use whenever the user shares text they suspect came from an LLM and wants it "humanized," "de-AI'd," "made to sound natural," or stripped of "AI slop" - em-dashes, smart quotes, pictographic characters, overused vocabulary ("delve," "leverage," "navigate the landscape"), hedging, and parallel three-item lists. For Russian text the equivalents are "уникальный," "не просто X, а Y," "не только... но и." Also triggers on requests to clean up LLM output, pass AI-detection tools, or fix text that "sounds like ChatGPT/Claude/Gemini." Do not just proofread; actively remove the patterns below.
---

# Humanize

Human prose has uneven sentence lengths, specific word
choices, and the occasional rough edge. AI prose has a
recognizable surface (certain characters, certain
vocabulary) and a recognizable rhythm (balanced clauses,
tricolons, every paragraph the same shape). This skill
covers both.

## STOP: pick the language track first

This skill contains two character tables that **contradict
each other**. The em-dash is the worst offender in English
and mandatory grammar in Russian. Straight quotes are
correct in English and wrong in Russian. Applying the wrong
table does not produce clumsy prose, it produces broken
prose.

So before reading any rule below, determine the language of
the text you are about to touch, and then read **only** its
track:

- **English text** -> [English track](#english-track).
  Skip the Russian track entirely.
- **Russian text** -> [Russian track](#russian-track).
  Skip the English track entirely.
- **Mixed document** (for example a chat log with Russian
  prose and English quotations): process each passage under
  the track for the language it is written in. Never apply
  one track's character rules across the whole file.
- **Any other language**: apply the shared rhythm rules
  only. Do not apply either character table or either
  vocabulary list, and tell the user which parts you
  skipped.

The [Rhythm and structure](#rhythm-and-structure) rules and
the [What to preserve](#what-to-preserve) rules are shared
and apply to every language.

## Two modes

**Drafting.** Apply everything below as you write. Do not
produce AI-sounding prose and clean it up afterward. The
lists here are things never to reach for in the first place.

**Rewriting.** Work in three passes, in this order:
characters, then vocabulary, then rhythm. Do not skip
ahead. Character fixes are cheap and mechanical. Rhythm
rewrites need judgment and should be done last, on text
that is already clean.

## English track

Everything from here to the Russian track applies to
**English text only**.

### Characters (English)

These are the clearest fingerprints of machine-generated
text. Avoid them when drafting; replace them globally when
rewriting.

| Avoid | Use instead |
|---|---|
| `—` em-dash (U+2014) | split into two sentences, or a colon `:` - rarely, a spaced hyphen ` - ` |
| `–` en-dash (U+2013) | hyphen `-`, or the word `to` in ranges |
| `"` `"` curly double quotes | `"` straight double quote |
| `'` `'` curly single quotes | `'` straight single quote |
| `…` ellipsis (U+2026) | `...` (three dots) |
| `•` `‣` `▪` bullet glyphs in prose | a real Markdown list, or remove |
| non-breaking space (U+00A0), zero-width space (U+200B), thin space (U+2009) | a regular space, or delete |
| decorative emoji and pictographs (✨🚀📊💡🎯) | remove, unless the user specifically asked for them |
| mathematical/fancy letters (𝐁𝐨𝐥𝐝, 𝘐𝘵𝘢𝘭𝘪𝘤) | normal letters; use Markdown `**bold**` / `*italic*` |

Em-dashes deserve special discipline. They are the single
most reliable AI tell. A human author uses at most one or
two per page; AI text often has one every paragraph. When
drafting, default to zero. When rewriting, prefer
restructuring over a literal hyphen substitution:

- Before: `The results were clear — the model had learned something real.`
- Good: `The results were clear. The model had learned something real.`
- Also good: `The results were clear: the model had learned something real.`
- Weak: `The results were clear - the model had learned something real.`
  (reads as AI with a find-replace applied)

### Vocabulary (English)

Plain English is almost always better. Entries marked
*(delete)* are usually filler that adds nothing once
removed.

**Verbs - pick a concrete one:**
- delve into -> look at, examine, get into
- leverage -> use
- utilize -> use
- foster -> build, encourage, create
- cultivate -> build, grow
- harness -> use, take advantage of
- unlock -> enable, allow, reveal *(often just delete)*
- empower -> let, help, allow
- embark on -> start, begin
- navigate -> handle, deal with, work through
- streamline -> simplify, speed up

**Adjectives - most are hedge-inflation:**
- robust -> strong, solid *(or delete)*
- seamless -> smooth *(usually delete)*
- comprehensive -> full, complete, thorough
- multifaceted -> complex *(or delete)*
- pivotal, crucial, paramount, vital -> important
  *(or delete)*
- cutting-edge, state-of-the-art -> new, current, advanced
- ever-evolving, ever-changing -> changing *(or delete)*
- transformative, groundbreaking -> new, big, major

**Metaphor nouns that appear in every AI paragraph:**
- landscape (of X) -> field, area, or just the noun itself
- tapestry -> *delete*, or "mix"
- realm -> area, field, world
- journey -> process, path
- treasure trove -> collection, set
- wealth of -> lot of, many
- myriad, plethora -> many, a lot of

**Phrases to never write:**
- "It's important to note that..." - just state the thing
- "It's worth mentioning..." - just mention it
- "In today's fast-paced world..." - delete
- "In the realm of X" - "In X"
- "At the end of the day..." - delete
- "In conclusion" / "In summary" - delete; if a summary is
  needed, write it without announcing it
- "As an AI language model..." - delete, obviously
- "I hope this helps!" - delete
- "Let's dive in" / "Let's explore" - delete
- "This is a testament to..." - rewrite directly

**Hedging inflation - tighten:**
- "could potentially" -> "could" or "might"
- "may possibly" -> "may" or "might"
- "it is generally considered that" -> "most people think"
  or just state the claim
- "some might argue" -> name who, or delete

## Russian track

Everything from here to the rhythm section applies to
**Russian text only**. Do not carry the English character
table into it.

### Characters (Russian)

The single most important rule: **the English em-dash rule
is inverted here.** In Russian the тире is not an AI tell,
it is compulsory grammar. Deleting it breaks the sentence.

**Keep these. Never strip them:**

| Character | Why it stays |
|---|---|
| `—` тире replacing an elided copula: `Москва — столица России`, `Наблюдать за торгами — отдельное удовольствие` | Mandatory. Russian has no present-tense "to be"; the dash *is* the verb |
| `—` тире for an elided repeated word: `верхней границей считается 1900 год, реже — 1875` | Mandatory |
| `—` тире opening a line of direct speech | Standard dialogue punctuation |
| `«` `»` ёлочки | The correct Russian quotation marks |
| `„` `“` лапки | Correct for a quote nested inside «...» |
| `–` en-dash in numeric ranges (`1941–1945`) | Normal Russian typography |
| `ё` | Leave every one in place. Do not "normalize" it to `е` |

**Replace these:**

| Avoid | Use instead |
|---|---|
| `"` straight double quotes | `«` `»` ёлочки |
| `"` `"` English curly quotes | `«` `»` ёлочки |
| decorative emoji and pictographs | remove, unless asked for |
| mathematical/fancy letters | normal letters; Markdown `**bold**` |
| non-breaking / zero-width spaces | a regular space, or delete |

`…` (U+2026) is fine: многоточие is ordinary Russian
punctuation, not a machine fingerprint. Normalize it to
`...` only if the user wants plain-ASCII output.

**Where the dash *is* a tell.** AI overuses тире as a
dramatic pause or a paired parenthetical aside, in
positions where a Russian writer would reach for commas,
brackets, or a full stop. Cut those; keep the grammatical
ones.

- Before: `Письмо — что особенно важно — было отправлено из Палермо.`
- After: `Письмо, что особенно важно, было отправлено из Палермо.`
- Before: `Ставка выросла — торги шли до последней секунды.`
- After: `Ставка выросла. Торги шли до последней секунды.`

### Vocabulary (Russian)

**Прилагательные — почти всегда инфляция:**
- уникальный -> особый, редкий, необычный *(чаще просто удалить)*
- ключевой -> главный, важный *(или удалить)*
- мощный -> сильный, большой
- инновационный, передовой, прорывной -> новый
- эффективный -> удалить или сказать, насколько
- комплексный, всесторонний -> полный
- глубокий (глубокий анализ) -> подробный
- значимый, существенный -> важный
- поистине, по-настоящему, действительно -> удалить
- невероятно, поразительно (как усилитель) -> удалить
  или поставить число

**Глаголы — берите конкретный:**
- погружаться (в тему) -> разбираться, изучать
- являться (как связка) -> это, или тире
- осуществлять, реализовывать -> делать, проводить
- обеспечивать -> давать, делать
- позволяет (X позволяет Y) -> пусть подлежащее
  действует само
- оптимизировать -> упростить, ускорить
- задействовать -> использовать

**Существительные-метафоры:**
- ландшафт (рынка, отрасли) -> рынок, отрасль
- экосистема (в переносном смысле) -> среда, набор
- палитра, спектр -> набор, ряд
- вызовы (калька с challenges) -> задачи, проблемы
- потенциал -> удалить или уточнить
- синергия -> удалить

**Фразы, которые не нужно писать:**
- "Важно отметить, что..." - просто скажите это
- "Стоит отметить / подчеркнуть..." - просто отметьте
- "Следует понимать, что..." - удалить
- "Таким образом," как автоматическая связка - удалить
- "В современном мире..." - удалить
- "Давайте разберёмся / рассмотрим подробнее" - удалить
- "В заключение", "Подводя итог" - удалить
- "Надеюсь, это поможет!" - удалить

**Канцелярит.** Nora Gal's diagnosis, and still the
loudest tell in Russian machine prose: stacked verbal
nouns and impersonal constructions. Turn them back into
verbs.

- Before: `Было принято решение об осуществлении проверки.`
- After: `Решили проверить.`

### Rhythm (Russian-specific)

The shared rhythm rules below all apply. Two contrast
patterns are specific to Russian and worth hunting
separately, because they are the local form of the
"not X, but Y" drumbeat:

- **`не просто X, а Y`** - AI reaches for this constantly.
  Assert the point directly instead.
  - Before: `Это не просто марка, а часть истории.`
  - After: `Эта марка — часть истории.`
- **`не только X, но и Y`** - same problem. Cut one half,
  or rewrite as a plain statement.

Also watch for every paragraph opening on a connector
(`Таким образом`, `Кроме того`, `Более того`, `При этом`).
One or two per page is normal; one per paragraph is a
machine.

## Rhythm and structure

These rules are **shared: they apply to every language.**
This is where AI prose still fails after the vocabulary is
clean. The words are fixed but the cadence gives it away.

**No tricolon habit.** AI loves three-item lists, especially
parallel ones: "clear, concise, and compelling"; "fast,
reliable, and scalable." A human uses tricolons sparingly.
Cut an item, change the structure, or replace the whole
list with one well-chosen word.

- Before: `The approach is fast, flexible, and effective.`
- After: `The approach is fast and it actually works.`

**No "not X, but Y" drumbeat.** Rhetorical contrasts like
"This isn't just a tool; it's a platform" get boring fast.
Prefer a direct assertion of the point you actually want to
make.

- Before: `It's not just about speed - it's about reliability.`
- After: `Reliability matters more than speed here.`

**Vary sentence length.** LLM prose clusters around 15 to
25 words per sentence, every one about the same size. Human
writing varies more. Mix long sentences with short ones. An
occasional fragment. Like this. Then a long, winding
sentence that carries the reader along and lets the
paragraph breathe before snapping back to something brief.
When rewriting, deliberately push some sentences under 8
words and others past 30.

**Vary paragraph shape.** Do not make every paragraph a
topic-sentence / two-supporting-sentences / wrap-up unit.
Let some paragraphs be a single sentence. Let others run
long without a tidy conclusion.

**Commit to a view.** When a real answer exists, take it.
Do not produce symmetrical "on one hand / on the other
hand" paragraphs on every question. Real writers have
opinions.

**No over-signposting.** Do not announce what the text is
about to do. Drop openers like "First, let's consider...",
"Now, turning to...", "To summarize...". Structure should
be obvious from the content, not from announcements.

## What to preserve

**Shared: applies to every language.** These are not AI
tells. Leave them alone.

- **Facts, numbers, names, and quotes.** Never smooth out a
  number or rename a person because it reads cleaner. If
  something seems factually wrong, flag it to the author.
  Do not silently edit it.
- **The author's argument and intent.** These rules govern
  surface and rhythm, not substance. Humanizing the voice
  must not change what the text says.
- **Technical terms in technical contexts.** "Leverage a
  load balancer" may be the right phrase in a systems
  chapter. Replace only when the word is filler, not when
  it is the accepted term of art.
- **Domain conventions.** Legal, medical, academic, and
  audit/compliance writing have their own rhythms that can
  look AI-ish but are not. Follow the conventions of the
  domain the text belongs to. Ask about context if unsure.
- **The other language's correct typography.** A Russian
  passage quoted inside an English document keeps its
  «ёлочки» and its grammatical тире. An English quotation
  inside a Russian document keeps its straight quotes.
  Punctuation belongs to the language of the passage, not
  to the language of the file.
- **Verbatim records.** In a transcript, chat log, or
  prompt archive, the human's own turns are evidence.
  Rewrite the machine's output; leave the human's words
  exactly as typed unless told otherwise.

## Delivery

Return the text itself. No preamble ("Here is the
chapter..."), no summary of the choices made, no closing
offer to revise. Default to clean output.

If the user asks what changed or why a word was picked,
explain then. Otherwise stay out of the way of the prose.

One exception, rewrite mode only: if the text runs past
roughly 500 words and you made many changes, you may add a
single closing line offering to flag them. Offer only. Do
not dump the diff unprompted.

## Example (English)

**Do not write this:**

```
In today's fast-paced digital landscape, it's crucial to
leverage cutting-edge tools to navigate the ever-evolving
challenges of remote work. This isn't just about
productivity - it's about fostering a seamless, robust, and
comprehensive workflow that empowers teams to thrive.
```

**Write this instead:**

```
Remote work keeps changing, and the tools for it keep
changing too. Picking the right ones matters. Not just for
getting more done, but for building a workflow the whole
team can actually rely on.
```

What changed: the opener cliche is gone, the tricolon is
gone, "leverage / navigate / foster / empower / seamless /
robust / comprehensive / cutting-edge / ever-evolving" are
all replaced or cut, sentence length varies, and no em-dash
was needed. The meaning survives. The AI voice does not.

## Example (Russian)

**Так писать не надо:**

```
Важно отметить, что данный экземпляр является поистине
уникальным объектом, обладающим ключевым значением для
коллекционеров. Это не просто марка, а настоящий символ
эпохи, позволяющий осуществить глубокое погружение в
ландшафт классической филателии. Таким образом, его
ценность невероятно высока.
```

**Надо так:**

```
Экземпляр редкий, и коллекционеры это знают. Марка 80
чентезимо на цельном письме встречается нечасто, а
вместе с погашенным штемпелем PD и американской
таксировкой — почти никогда. Отсюда и цена.
```

Что изменилось: ушли "важно отметить", "является",
"поистине", "уникальный", "ключевой", "погружение",
"ландшафт", "не просто X, а Y" и "таким образом";
вместо оценки "невероятно высока" появились конкретные
факты. Тире осталось там, где оно требуется грамматикой.
Смысл сохранён, машинный голос исчез.

---
name: design-doc
description: Create and iteratively evolve a software architecture design document — formal spec first, then Q&A-driven refinement, a consolidation pass, an optional story-style rewrite for human readability, and polished PDF output. Carries the durable principles (simplicity, modularity, self-healing, continuous verification, supply-chain discipline) that every such document should state. Use when the user asks to create, review, improve, or rewrite a design document, architecture document, or technical proposal.
---

# Creating a Software Architecture Design Document

Guides an iterative, conversation-driven process for producing a design
document that is technically rigorous **and** genuinely readable.
Distilled from a real project (a knowledge base over millions of
documents) that went from formal spec to story-style document over ~30
rounds of dialogue. The examples are from that project; the method
applies to any system.

**Core insight:** a design document is not written in one pass. It
evolves through `question → clarification → fold the clarification back
into the document`. The reader's questions are the best signal for what
the document fails to explain.

**Scope check first.** A PRD (Product Requirements Doc) says *what* and
*why*, written by product for everyone. An ADD (Architecture Design Doc)
says *how* and *why that way*, written by engineers for engineers, with
the PRD as input: `PRD → architecture design → implementation`. This
skill is about the ADD. If no PRD exists, extract requirements from the
requester and put them in an appendix — a design that can't point at the
requirement it satisfies is an opinion with diagrams.

## Ground Rules (every phase)

1. **Propose before editing.** "Does it make sense to…" / "can we
   improve…" wants an assessment, not a rewrite. Present proposals
   ranked by value against cost; edit only when told ("OK, update the
   document").
2. **Answer questions in plain language first.** When the user is
   confused ("is the wiki in files or in the database?"), answer
   conversationally with analogies and concrete examples. *Then* ask
   whether the confusion reveals a gap worth fixing in the document.
3. **Preserve before destructive rewrites.** Before restructuring
   (formal → story), commit to git or save the current version under a
   distinct name (`<name>-formal.md`). Tell the user where it is.
4. **Verify claims against the actual text.** "Confirm the design does
   X" means read it and cite sections. If it doesn't do X, say so.
5. **Fact-check external references.** If the user names a format,
   standard, or project you can't verify, flag it honestly and propose
   the real established equivalent.
6. **Markdown is the source of truth.** PDFs, HTML, and rendered
   diagrams are projections — regenerate, never hand-edit.

## The Principles Catalog

Make sure the document states these, adapted to the domain. They are
the durable core that transfers to any system.

1. **Simplicity is the guiding principle.** The simplest architecture
   that still delivers the *full* functionality — the fewest moving
   parts, each chosen deliberately. Simplicity makes a project do-able,
   flexible, maintainable, affordable, fast to prototype: no big team,
   no fleet of clusters, no heavy hardware. Complexity in architecture
   and infrastructure can bring a project to its knees — projects rarely
   die from missing features, they drown in their own plumbing. Prefer
   one well-understood system over several specialized ones; prefer a
   SQL join over a new service. *Complexity must justify itself with
   evidence; simplicity is the default.* Not minimalism: never sacrifice
   required functionality, correctness, security, or auditability.
2. **Bets are falsifiable — write scaling gates.** Each simplification
   gets a measurable threshold at which you'd change your mind ("chunk
   count above 50M; p95 above 2s despite tuning"). Turns "will it
   scale?" into "which gate are we near?"
3. **Modularity — don't let the dumplings fuse.** Simplicity's evil twin
   is the small system nobody can change. Each module: a wrapper keeping
   its filling to itself (encapsulation), its own distinct filling
   (separation of concerns / bounded context), cookable and tastable
   alone (developed and tested separately), touching neighbors without
   sticking (loose coupling). The failure mode is shortcuts under
   deadline — one component queries another's tables, a parser learns
   what's downstream — until the diagram still shows few boxes but
   nothing can change without understanding everything. Requires:
   services not schemas; external systems as plugins behind normalized
   events; implementations behind contracts; **dependencies one-way**
   (domain ← application ← adapters ← bootstrap); explicit wiring in one
   composition root (no service locators, import-time registration,
   global mutable state, monkey-patching). Test for every boundary: *can
   this be changed, replaced, or removed without understanding the whole
   system?*
4. **Modularity reaches into the code.** Put these in the document, not
   just a style guide: files **under 800 lines**; functions **under 50
   lines**; a doc block at the top of every file and on every function
   and class; code split into subdirectories along module boundaries
   with **a `README.md` per directory** stating what it owns, its public
   API, and its dependencies.
5. **One system of record.** Name the single place truth lives.
   Everything else is a copy, cache, or projection — and **projections
   are never hand-edited**; they render from the source and can be
   deleted and regenerated.
6. **Do the expensive work once; everything downstream is disposable.**
   Identify the step costing weeks of compute (parsing, training,
   enrichment); store its output permanently; everything after reads
   from *that*. This is what makes future migrations boring.
7. **Boring technology — and justify the language.** Someone will ask
   "why not rewrite it in <faster language>?" Answer it in the document:
   where does time actually go (if hot paths are already C/CUDA/the
   database, rewriting glue speeds up the glue, and the glue was never
   the bottleneck); where does the ecosystem live; who maintains it; and
   note that modularity keeps it reversible — a measured CPU-bound
   module can be swapped behind the same contract later.
8. **Trust is re-earned continuously** — see [Verification](#verification-two-timescales).
9. **Self-healing.** At scale something is always slightly broken. If
   every small failure needs a human, on-call becomes the immune system
   — and humans make terrible white blood cells. So: crashed worker →
   idempotent retry from a transactional queue; corrupted file →
   quarantine, restore from the most recent *verifying* snapshot,
   re-check, file a report the human reads after the repair; missed
   event → reconciliation queues the refetch (detection and repair are
   one motion); degraded index → automatic rebuild in the next
   maintenance window; anything derived → throw it away and regenerate.
   **Manners:** backoff, finite retries, dead-letter + page (infinite
   retry is not persistence, it's a tantrum). Name what is **never**
   auto-repaired — typically permissions, money, regulated records. *The
   system heals itself; it does not improvise itself.*
10. **Don't chase the latest version.** The freshest package is the
    least reviewed package; "latest" is not a version, it's a gamble.
    Every dependency is pinned, aged, verified: runtime from a pinned
    major.minor (`brew install python@3.12`); packages via `uv` with
    `exclude-newer = "30 days"` and `prerelease = "disallow"` in
    `pyproject.toml`, lockfile committed, upgrades an explicit act
    (`uv lock --upgrade && uv sync`); databases on a pinned major from
    the official repo, never a fresh `.0` in production; container
    images official-only, **pinned by digest** (`:latest` banned),
    pulled once into an internal registry and scanned on the way in.
11. **Don't forget the humans.** A system without windows gets operated
    by SSH — an incident waiting for a typo. Include a **command
    center** (health, lag, quality trends, reconciliation deltas, what
    self-healing did overnight; plus start/pause/re-run controls), an
    **operator chat** if agents are in play, and the **end-user
    surface** with one-click flagging of bad results into the eval queue
    — the complaint department feeds the immune system. Two rules: every
    button drives the same documented, audited APIs (a dashboard is a
    view with steering, not a backdoor), and state-changing actions
    require confirmation and land in the audit log. Frontend house
    style: **vanilla JavaScript, no frameworks, no build step**; small ES
    modules, same size/doc limits as the backend; **all styles in one
    `styles.css`**.
12. **Give the work an owner.** Maintenance work needs owners, statuses,
    and handoffs, or it lives in email threads and in the heads of
    people who might be on a beach. Add a maintainer task system (the
    reference project calls a task a "monkey" — the next move). The
    tasks already exist: exhausted retries → dead-letter *becomes a
    task*; eval gate trips → task with the regression report; user flags
    a result → task with the trace; scrub/reconciliation failures →
    tasks. The immune system files tickets. Mechanics: each person sees
    their own; vacation transfers the whole set in one action; a group
    dashboard shows who's drowning. Keep it a **separate app** so the
    core service stays lean while the workbench evolves at business
    speed. Maintainers may **vibe-code** their own tools — with
    guardrails, not a leash: inside the maintainers' app, only through
    documented APIs (so ACLs and audit apply automatically), passing the
    same CI, size limits, and conformance review.

## Verification (two timescales)

Both failure modes are silent: code changes with no compiler complaint
when someone imports infrastructure into the domain, and data changes
that degrade quality with no error anywhere. *Silence is just silence.*

**Guarding the code — a pyramid, all four layers:**

- **Unit** — functions; fast, isolated, no network or DB, dependencies
  faked at the port interfaces (only possible because ports exist).
- **Module** — each module tested through its **public API**,
  dependencies stubbed. If a module's tests can't be written without
  reaching into another module's internals, a boundary is wrong.
- **Integration** — real database in a disposable container (pinned by
  digest); pipeline end-to-end against a fixture dataset.
- **Architecture conformance** — mechanical rules to mechanical tools
  (`import-linter`, file-length / function-length / docstring checks in
  CI); judgment rules to **an AI agent that reviews every change against
  the written architecture rules**, flags the specific rule it believes
  is broken, and gates the merge. Humans adjudicate; the agent ensures
  erosion never happens *silently*.

**Guarding the behavior — evals grown from real outputs.** The tempting
shortcut is a conference-room checklist scored 1–10. It fails twice: the
scores are too mushy to act on, and the criteria only cover predicted
failures. So:

- Start from **real outputs and traces**, not an abstract checklist.
- **Review failures before writing criteria** — experts annotate
  problems *in context*.
- Split criteria: **top-down** (what the task always required — format,
  length, policy, citation present) and **bottom-up** (recurring flaws
  found comparing real outputs against human-approved ones).
- **Specific pass/fail checks, never vague scores.** Not "faithfulness:
  7/10" but "every numeric claim appears in a cited source: yes/no."
  Binary checks are debuggable.
- **One judge pass per criterion** — one mega-prompt verifying twelve
  requirements will quietly forget a few.
- **AI does the labor, humans supply the taste** — AI clusters failures,
  surfaces patterns, builds the review interface, drafts rubric items;
  only the domain expert knows what's *materially* wrong.
- **Graduate every validated pattern** into a regression case, a CI
  check, a dashboard series, a production monitor.
- **Gate on it:** nightly run vs. a rolling baseline; regression beyond
  threshold pages on-call and **freezes updates**. A failed eval is
  treated exactly like a failed deploy.
- Add **reconciliation counts** across layers (source → registry →
  serving) so a dead connector is a mismatched number tonight, and
  **self-retrieval probes** proving new content is findable, not merely
  stored somewhere warm.

One principle at both timescales: **a failing test blocks the merge; a
failing eval freezes the pipeline.**

## Phase 1 — The Formal Draft

Get the substance down; ugly is fine, complete matters more. Section
order (adapt; cut what doesn't apply — it's the order a reader needs to
learn things):

1. **Title + status line** (`Status: Draft (pending approval) · Date: …`).
   No version numbers in file names, no changelog at the top — git is
   the changelog.
2. **Executive summary** — 3–5 short paragraphs: scale headline, key
   bets, resource footprint. Write it last, place it first.
3. **Table of contents** with working anchors.
4. **The problem** — what's broken today, concretely; then the fine
   print constraints (compliance, scale, latency, cost, residency).
5. **Explicit non-goals** — what v1 deliberately does *not* do. The
   cheapest scope protection there is.
6. **The simplicity bet** — the architecture, why this few parts, the
   scaling gates, the language justification.
7. **Modularity** — seams, contracts, dependency direction, code rules.
8. **Storage and data model** — with actual DDL/schemas, plus durability
   disciplines (content addressing, write-once) where relevant.
9. **Processing pipeline** — with the "expensive work once" boundary
   marked.
10. **Core domain sections** — three to six, specific to the system.
11. **Lifecycle** — add / update / remove propagation.
12. **Verification** — test pyramid + evaluation gate.
13. **Self-healing** — reflexes and their limits.
14. **Correctness contract** — what the output guarantees, and what the
    system does when it doesn't know.
15. **Human interfaces.** 16. **Task ownership.** 17. **Disaster
    recovery.** 18. **Sizing.** 19. **Dependency discipline.**
20. **Roadmap** — phases with timeframes; trust-building work
    (evaluation, sign-offs, DR drills) scheduled *before* scale-out.
21. **The shape of the thing** — the design distilled to 6–8 numbered
    decisions, simplicity first. Executives read this section; make it
    the best writing in the document.

**Appendices:** A — Requirements with stable IDs (Functional `F1…`,
Security/Compliance `S1…`, Non-Functional `N1…` with *measurable*
targets); B — Technology choices (layer / selection / why, written for
someone who will challenge it); C — Risk register (risk / engineered
mitigation), condensed.

**Diagrams** (SVG, so GitHub and Obsidian render them): architecture,
data layout, main workflow, roadmap. Four good ones beat forty.

### The sections everyone forgets

Formal drafts describe systems *at rest*. Reality is motion. Write these
before a reviewer has to ask:

- **Lifecycle.** Every derived artifact updated on add / update /
  remove. Include a **propagation matrix** (event × artifact → mechanism
  and latency) and the hard rules: **soft-deactivation not deletion**
  where history must stay citable; **atomic visibility flips** (old off,
  new on, one transaction — no query catches a record half-dressed); and
  **permission changes outrank everything** (stale content is
  embarrassing, stale permissions are a breach) with their own
  minutes-scale SLA, their own alert, and a daily full re-sync.
- **Continuous evaluation** — its own section, not a paragraph.
- **Self-healing** — its own section.
- **Backup and recovery.** Open with the unpopular truth: **replicas are
  not backups** — a standby faithfully replays your `DROP TABLE` in
  milliseconds. Specify data tiers (irreplaceable vs. reconstructible),
  RPO/RTO as numbers, backups out of blast radius (separate account,
  independent credentials, snapshot locks, WORM), **restore *ordering***
  (not obvious under pressure — write it down), and **scheduled restore
  tests**, because a backup never restored is a rumor: monthly automated
  restore + verification, quarterly timed drill, failed restore pages
  like an outage.
- **Sizing with arithmetic shown** (`3M docs × ~1 MB ≈ 3 TB`). Tables
  for storage and database, steady-state total, provisioning with
  headroom. Sizing is what makes a simplicity bet look rational rather
  than reckless.
- **Human interfaces and task ownership.**

## Phase 2 — Interrogate It

Expect these moves; each has a right response. The bracketed phrasings
are reusable prompts.

- **"Confirm the design does X — but don't change anything, only propose
  improvements."** Verify against the text, cite sections. That
  don't-change clause is the most valuable half-sentence in the process;
  honor it strictly.
- **"Would it make sense to add Y?"** Assess honestly: already covered
  (say where), real gap, or marginal. Rank by value/cost. Wait for the
  go-ahead.
- **"So we will have A, B, C, D — correct?"** Confirm with a mapping
  table: user's term → design's term → status. Clarify relationships
  between layers the user sees as separate.
- **"I'm confused about Z — is it in files or in the database?"**
  Confusion is diagnostic. Answer plainly, then fix the document.
- **"Can we borrow from project Z?"** (URL) → fetch and study it, then
  three buckets: *already covered*, *worth adopting* (concrete schema/API
  changes), *marginal*. Note what doesn't transfer.
- **Gap-finding asks** — describe the update process when data is
  added/removed/updated including every derived artifact; describe how
  accuracy and completeness are evaluated after daily updates; does the
  design include an interface for humans; does it prescribe unit,
  module, integration, and AI-agent architecture-conformance tests.
- **Principle injections** — the user names a principle, states the
  reasoning in their own words, and asks for it in problem→solution
  form. Deliver it that way; principles inserted like this land as
  arguments the reader can follow, not slogans.
- **"Update the document per our conversation"** → integrate everything
  consistently: requirement rows, principles, sections, tech table,
  roadmap, risk matrix, cross-references. Renumber carefully and fix
  every reference to renumbered sections.

## Phase 3 — Consolidate

**Not optional.** Run whenever the document has grown ~20% since the
last pass: *"We were adding topics incrementally and the document kept
growing. Review the whole document to see if it can be rearranged or
rewritten to be shorter and more elegant."* Documents grown by accretion
always carry duplication their author can't see.

Also in this pass: executive summary at top; changelog and filename
version numbers removed; TOC anchors verified in both GitHub and the PDF
build; no dangling references to removed material.

## Phase 4 — Story Rewrite (offer when the formal draft is content-complete)

Same technical substance, told as challenges converted into solutions.
Nothing is dumbed down — every number, threshold, and trade-off stays;
only the order of presentation changes, from taxonomy to narrative.

**Each section is one concrete challenge:** open with the problem
vividly (a scene, a failing query, an uncomfortable truth) → explain why
the obvious answers fail (this pre-empts the reviewer about to suggest
exactly that) → deliver the solution in a **bolded sentence** so a
skimmer still gets the architecture → keep the rigor → close by
connecting forward or back.

Style rules:

- **No "Prologue", "Chapter N", "Epilogue".** Plain evocative headings:
  *"The Day Everything Goes Wrong"*, *"What Was True Last March?"* A
  good heading is the section's question.
- **Cross-reference by name or idea, never by number** — numbers break
  when sections move.
- **Concrete numbers over adjectives.** Not "highly scalable" but "25
  million chunks, which for a well-fed PostgreSQL instance is a
  Tuesday."
- **Short declarative sentences for the punches.** *"Nobody can find
  it."* *"Silence is just silence."*
- **Explain *why* before *what*.**
- **Tables only for genuinely tabular data**; prose for reasoning.
- **Light, not cute.** Wit keeps a forty-page document alive, but the
  joke must carry the point. If a sentence is only funny, cut it.
- **Use words the readers know.** "Avoiding the Pelmeni Architecture"
  became "Don't Let the Dumplings Fuse" — and the metaphor got sharper.
  **If a metaphor needs a footnote, replace it.**
- **Close with "the whole design is N decisions."**
- **Formal material becomes appendices.**

Keeping a formal version alongside is optional — one document everyone
reads beats two nobody finishes.

## Phase 5 — PDF Generation

Target US Letter, 0.7" margins (or as requested). No LaTeX needed:

1. `pandoc -f gfm -t html5 --standalone --css=<style.css> design.md -o build/design.html`
2. **WeasyPrint** → PDF.
3. CSS: `@page { size: Letter; margin: 0.7in; }` + page counters in
   margin boxes; running header suppressed on page 1;
   `header#title-block-header { display: none; }` (pandoc duplicates the
   title); `tr { page-break-inside: avoid; }`; `page-break-after: avoid`
   on headings; wrapped `pre` blocks; compact styled tables.
4. **Rasterize SVG diagrams** — WeasyPrint mangles SVGs with embedded
   font styles. Headless Chrome per diagram:
   `--headless --screenshot=out.png --window-size=1400,900 --force-device-scale-factor=4`.
   Swap `.svg` → `.png` with `sed` **at build time only**; the Markdown
   keeps SVGs.
5. **Verify:** page count and size; visually inspect the title page, a
   table-heavy page, and every figure page; confirm TOC anchors resolve
   (`href="#…"` vs `id="…"` in the built HTML); grep the extracted PDF
   text for newly added sections.
6. Keep the CSS in the repo (`.pdf-style.css`), document the rebuild
   one-liner, and **regenerate after every content change** — say "and
   the PDF" or it silently drifts out of sync.

Also fix mechanical damage found along the way: corrupted escape
sequences (`\r`/`\t` eaten from `$\rightarrow$`), image paths pointing
at directories that don't exist, anchors broken by renamed headings.

## Anti-Patterns

| Anti-pattern | Fix |
|---|---|
| **The vendor stack** — five specialized systems because every reference architecture has five | Do the arithmetic, write the scaling gate, start with one |
| **The option menu** — "we could use A, B, or C…" | Decide. State the reason. Move on |
| **Adjective architecture** — "highly scalable, robust, cloud-native" | Replace every adjective with a number |
| **Steady-state only** — no word on what happens when data changes | Write the lifecycle section |
| **The frozen checklist** — criteria invented in a conference room, scored 1–10 | Grow criteria from real traces; binary pass/fail |
| **`design-v4-FINAL-v2.md`** | Git is the changelog. One filename, forever |
| **The changelog header** | Delete it. Git remembers |
| **Accretion bloat** | Run the consolidation pass every ~20% growth |
| **The unread masterpiece** — technically perfect, nobody finishes it | Story rewrite |
| **Machines only** — no dashboard, no chat, no task list | Human interfaces; give the work an owner |
| **The footnoted metaphor** | If it needs a footnote, replace it |
| **Stale PDF** | Regenerate every time |

## Definition of Done

- Scale, cost, and "how big is this really?" answered **with numbers**,
  arithmetic shown.
- **Simplicity** stated as a principle; every component earns its place
  against it; every bet has a **written scaling gate**.
- **Modularity** specified at both levels — contracts between modules,
  size and doc limits inside them.
- **Lifecycle** (add/update/remove propagation) described, not just
  steady-state architecture.
- **Verification** covers both timescales: a test pyramid gating merges,
  an evaluation gate gating updates.
- **Self-healing** reflexes listed, with what is never auto-repaired.
- **Backup and recovery** states RPO/RTO, restore ordering, and a
  restore-testing schedule.
- **Humans** have interfaces; maintenance work has owners.
- **Dependencies** pinned, aged, verified. **Non-goals** explicit.
- Closes with **N decisions** anyone can quote.
- A newcomer reads it in one sitting; an implementer builds from it plus
  the appendices.
- The PDF is regenerated, verified, and matches the Markdown.

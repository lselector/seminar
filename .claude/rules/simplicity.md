# Simplicity

When designing architectures, planning code structure, or proposing solutions,
always use simplicity as the guiding principle.
Find simple, elegant, minimal solutions. Always try to simplify.

## The Principle

Choose the simplest architecture that still delivers the full functionality —
the minimum number of moving parts, chosen deliberately.

Simplicity makes a project:
- **Do-able** — it can actually be built and shipped
- **Flexible** — fewer parts to rearrange when requirements change
- **Maintainable** — fewer things that can break, fewer experts required
- **Affordable** — no fleet of clusters, no heavy hardware, no big team
- **Fast to prototype** — something real in front of users in weeks, not quarters

Whereas complexity in architecture and infrastructure can bring a project to its knees.
Every additional system needs specialists to run it, integration code to connect it,
and attention when it breaks. Projects rarely die from missing features;
they die from drowning in their own plumbing.

## How to Apply

- Prefer one well-understood system over several specialized ones.
  Add a new component only when a measured, written-down threshold demands it
  (a "scaling gate"), not because a vendor pitch or fashion suggests it.
- Prefer boring, proven technology the team already knows.
- Prefer a SQL join, a plain file, or a library call over a new service.
- When two designs deliver the same functionality, choose the one
  with fewer moving parts, fewer dependencies, and less code.
- When writing code: fewer abstractions, fewer layers, fewer options.
  Do not add flexibility for imagined future needs (YAGNI).
- Actively look for simplifications in existing designs and code —
  removing a component is a win worth proposing.
- Complexity must justify itself with evidence; simplicity is the default
  and needs no justification.

## What This Does NOT Mean

- Do not sacrifice required functionality, correctness, security, or auditability in the name of simplicity. 
 The goal is the *simplest solution that fully works* —
  not the simplest solution.

- Do not conflate "simple" with "quick hack". Simple designs are usually the result of more thought, not less.

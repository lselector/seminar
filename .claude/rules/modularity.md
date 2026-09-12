# Claude Code Rule: Modular, Plugin-Oriented Software

Build modules like good dumplings on a plate: a wrapper that keeps the filling in (encapsulation), one distinct filling (a single bounded responsibility), cookable and tastable alone (developed and tested separately), touching its neighbors without sticking (loose coupling). The failure mode is dumplings fusing into one clump through hidden imports, shared state, and cross-module shortcuts — then nothing can be changed or tested separately.

## Core Rules

1. **Choose ownership first.** Name the module that owns the feature, its public API, and its allowed dependencies. Other modules use only that documented API — never another module's internals, private files, persistence, or mutable state.

2. **Depend on ports, not vendors.** Core code depends on small interfaces; databases, SDKs, HTTP clients, queues, file systems, and LLM providers live behind them in replaceable plugin/adapter modules. Keep ORM entities, framework request objects, and vendor types out of public contracts — use typed inputs/outputs with defined errors.

3. **Keep dependencies one-way and acyclic.** `domain` (rules, entities, ports; no framework or infra) ← `application` (use cases) ← `plugins` (implement ports) ← `interfaces`/`bootstrap` (config and wiring only). Fix a cycle by moving a contract down, extracting an abstraction, or using an event — never with late/dynamic imports.

4. **Wire explicitly.** Construct concrete plugins only in a small composition root; pass them via constructors or arguments. No service locators, global mutable state, import-time registration, implicit singletons, or monkey-patching.

5. **Stay cohesive.** One responsibility per module, organized by feature/domain. No catch-all `utils`, `common`, `helpers`, `shared`, or `models`.

6. **Test in isolation.** Unit-test every module with fakes for its ports, integration-test adapters, and reserve end-to-end tests for critical workflows.

## Workflow

Before implementing, state the owning module and its public API, any new or changed port contracts, the dependency direction, and the tests to add.

A change is done when dependency flow is acyclic and inward, external systems sit behind adapters, a concrete plugin can be swapped without touching core logic, module tests run in isolation, and any exception is documented with its reason, scope, and containment plan.

**Principle:** A piece should be changeable, replaceable, testable, or removable without understanding or modifying the whole application.

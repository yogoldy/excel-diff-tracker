# Product strategy map

This folder preserves product reasoning for Excel Scenario Analysis Tool in a form that is useful to
both the maintainer and future agents. It is intentionally separate from implemented-state
documentation: [Project status](../PROJECT_STATUS.md) records current maturity,
[Product roadmap](../PRODUCT_ROADMAP.md) records provisional delivery direction, and
[Architecture](../ARCHITECTURE.md) describes only implemented mechanics.

## Start here

- **Five-minute orientation:** [Product thesis](PRODUCT_THESIS.md)
- **What the possible products are:** [Product layers](PRODUCT_LAYERS.md)
- **What the software may safely read and write:** [Trust and execution boundary](TRUST_AND_EXECUTION_BOUNDARY.md)
- **Who might buy it and what already exists:** [Market and competition](MARKET_AND_COMPETITION.md)
- **Why these ideas emerged and which nuances remain unresolved:**
  [2026-09-03 discovery synthesis](sessions/2026-09-03-scenario-product-discovery.md)
- **Current delivery sequence:** [Product roadmap](../PRODUCT_ROADMAP.md)

## Status vocabulary

These notes deliberately distinguish:

- **Implemented:** present in a tested build and documented in Project Status or Architecture.
- **Planned:** provisionally accepted as a direction, but not implemented.
- **Hypothesis:** plausible product, market, or design proposition requiring validation.
- **Research target:** a competitor, technology, or user behavior to inspect before deciding.

Nothing in this folder expands the capabilities claimed for version 0.2.0.

## Using this folder with Obsidian

The repository can be opened directly as an Obsidian vault. No Obsidian installation, plugin, or
tracked `.obsidian` configuration is required. Standard Markdown links are used so the same notes
remain navigable in GitHub, Codex, ordinary editors, and Obsidian.

The map above is the human entry point. Session notes preserve reasoning and unresolved nuance;
focused notes preserve the current synthesized view. When a hypothesis becomes an accepted roadmap
decision, update the focused note and Product Roadmap rather than relying on a buried session passage.

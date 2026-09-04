# Trust and execution boundary

The product-design choice is not a binary choice between read-only access and unrestricted workbook
editing. Three capabilities must remain distinct.

## Observe

The application may read workbook structure and saved values, record chronological snapshots,
compare baselines and scenarios, and export evidence. This is the implemented foundation in 0.2.0,
subject to the exact limits in Project Status and Architecture.

## Operate

A future controlled runner may write values only into analyst-approved input cells. It must not
decide which cells are inputs at runtime. The analyst-approved scenario contract defines addresses,
labels, types, ranges, permitted values, outputs, and validations.

An intended run would behave like a transaction:

1. Verify the workbook identity and compatible version.
2. Confirm the expected formulas or structural fingerprint around the scenario contract.
3. Use an isolated working copy or controlled Excel session rather than casually changing the
   canonical file.
4. Snapshot every approved input value.
5. Write only to an explicit allowlist.
6. Validate type, unit, range, and permitted values.
7. Recalculate using real desktop Excel when authoritative runtime calculation is required.
8. Capture approved outputs, whole-workbook differences, calculation state, and validation results.
9. Restore the prior values or discard the working copy.
10. Save the scenario record separately from the workbook.

This behavior is planned, not implemented in 0.2.0.

## Build

Changing formulas, inventing dependencies, restructuring sheets, or creating missing economic logic
is model development. It should not be part of an unattended or management-operated scenario run.
Such work requires explicit analyst review and a separate validation process.

## Deterministic runtime and optional AI

AI can be valuable at design time under analyst oversight. It may help identify candidate levers,
group changed outputs, explain formulas, draft documentation, or propose validation checks. It must
not become the numerical authority.

The intended trust model is:

```text
Design time: analyst judgment, optionally assisted by AI
Runtime: approved contract + deterministic writes + Excel calculation + fixed controls
Presentation: structured outputs; optional narrative constrained by captured evidence
```

Management should operate a preconfigured scenario, not ask an AI agent to reinterpret a sensitive
workbook from scratch. Core reports and exports should remain available with AI disabled.

## Privacy

Local-only operation is both a product feature and a trust boundary. Sensitive workbook contents,
scenario values, and calculation results should not require server hosting or network transfer. Any
future optional network or AI capability must be explicit, separately consented, and unnecessary for
deterministic scenario execution.

## Failure posture

The controlled runner should fail closed when:

- the workbook identity is unexpected;
- a mapped sheet or cell no longer exists;
- a protected or formula cell would be overwritten unexpectedly;
- two mutually exclusive scenarios are active;
- Excel calculation does not complete successfully;
- validation or reconciliation fails; or
- restoration cannot be verified.

A failed run may retain diagnostic evidence, but it must not be represented as a certified scenario.

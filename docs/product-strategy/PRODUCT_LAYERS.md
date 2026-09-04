# Product layers

The product can deliver value progressively. A customer should not need sophisticated governance or
a purpose-built scenario engine before receiving a useful result.

## Implemented foundation: workbook diff tracker

Version 0.2.0 already provides local save-by-save workbook capture, immutable chronological deltas,
selected-baseline comparisons, and separate direct-value, formula-text, and calculated-result
changes. See [Project status](../PROJECT_STATUS.md) for exact maturity and qualification limits.

This foundation is mechanical evidence. It does not yet create, execute, or certify scenarios.

## Layer 1 hypothesis: scenario recorder

The user begins at a known saved state, names a scenario, optionally identifies lever and output
cells, edits the workbook normally, and finishes recording. The product preserves the complete
workbook diff as a safety net and exports a structured scenario record.

This mode can work even when a workbook has limited scenario architecture because the user performs
the edits in Excel and the tracker records what happened.

## Layer 2 hypothesis: controlled what-if runner

For a workbook with existing calculation logic, the analyst maps approved inputs and outputs. The
software writes values only to approved input cells in an isolated working copy or session, asks
desktop Excel to recalculate, captures outputs, restores or discards the changed state, and records
the complete run.

This is likely the strongest initial standalone product because it converts an informal process into
a reusable one without asking the software to understand or rewrite the full model.

Expected capabilities include:

- multi-sheet input and output selection;
- human-readable labels, units, and value validation;
- automatic baseline capture and restoration;
- named, duplicable, and comparable scenarios;
- workbook-version compatibility checks;
- formula-cell warnings for selected inputs;
- output watchlists and materiality thresholds;
- complete local run logs; and
- structured export to GCBS or ordinary files.

## Layer 3 hypothesis: scenario assurance

The assurance layer adds controls needed for consequential models:

- analyst certification of a scenario contract;
- inactive-scenario leakage detection;
- recurring, one-time, and lagged timing classifications;
- evidence-derived values versus management assumptions;
- reconciliation against authoritative workbook outputs;
- expected invariants and validation cells;
- leadership-facing explanations backed by analyst-level audit detail; and
- controlled evolution when workbook sheets, ranges, or formulas change.

## Layer 4 hypothesis: team and portfolio system

A later corporate product could support shared scenario libraries, approvals, permissions, model
registries, explicit combination of independently validated initiatives, and organization-wide
decision packages. Combined scenarios must be intentional and checked for interaction or double
counting; they must not arise from accidental addition of live inputs.

## Product surfaces

Two surfaces may use the same underlying scenario record:

1. **A simple Excel scenario-manager add-in** for selecting cells, saving cases across multiple
   worksheets, comparing results, restoring baseline, and exporting. This may be a low-cost,
   search-led utility.
2. **A corporate scenario-assurance application** for certification, validation, evidence,
   collaboration, and management-ready operation.

The small add-in is a distribution and onboarding wedge, not a disposable prototype. Every scenario
it records can contribute to the richer product's reusable contract and evidence model.

## Applicability boundary

A workbook is immediately suitable for the controlled runner when it already has identifiable input
cells, formulas that propagate those inputs, and observable outputs. A static statement made mostly
of pasted values can still be recorded and compared, but the product must not claim it contains
scenario logic that is not present.

Creating missing driver logic is a separate model-development activity. It may be supported by
analysts, consultants, templates, or supervised AI, but it is not the trusted scenario-runner path.

# Product roadmap

This document owns future product direction for Excel Scenario Analysis Tool. It distinguishes the
implemented foundation from planned workflows and hypotheses. Current maturity belongs in
[Project status](PROJECT_STATUS.md); implemented internals belong in [Architecture](ARCHITECTURE.md).

The supporting [product strategy map](product-strategy/README.md) preserves the product thesis,
layered possibilities, trust boundaries, market research, and discovery-session reasoning behind
this roadmap. Those notes do not establish implementation status.

## Product goal

Turn a general Excel workbook into a reusable scenario-analysis instrument without requiring the
application to understand one particular model. The product should preserve exact saved evidence,
help a user identify levers and outputs, and make chosen scenarios available to downstream
presentation tools.

## Implemented foundation in 0.2.0

- Local save-by-save workbook capture and immutable chronological deltas.
- Separate direct-value, formula-text, and calculated-result changes.
- Comparison against the previous save, original scan, or a selected saved scan.
- Derived comparison exports that do not rewrite chronological evidence.
- Compact and expanded workbook/history presentations with path actions.
- Local-only storage with no account, telemetry, or workbook-content network transfer.

These capabilities are mechanical evidence. They do not determine intent, business correctness, or
which changes matter to a decision.

## Planned scenario workflow

The target workflow is deliberately simple:

1. Start recording from a known saved workbook state.
2. Name the scenario manually or bind its name to a selected cell.
3. Optionally designate lever cells that the user intends to change.
4. Optionally designate output cells or ranges that express the decision result.
5. Make and save changes normally in Excel.
6. Finish recording and review the complete saved-state comparison.
7. Preserve the named scenario for later comparison or downstream visualization.

Selected levers and outputs organize the scenario; they are not the detection boundary. The
whole-workbook diff remains the safety net so unselected changes are still visible. User-provided
intent remains authoritative, while the tracker supplies corroborating mechanical evidence.

## Provisional product stack

- **Desktop host:** owns local history, saved comparisons, named scenario records, evidence, and
  exports.
- **Thin Excel add-in:** provides in-workbook start/finish controls, scenario naming, and range
  selection. It should send references and user intent to the local host rather than duplicate the
  history engine.
- **Downstream visualization:** tools such as GCBS consume deliberately exported scenario definitions
  and outputs for comparison and presentation. They do not replace the evidence host.

This division is a hypothesis pending specification and testing. The add-in and downstream contract
are not implemented in 0.2.0.

## Optional semantic mapping

Model-specific maps may label controls, behavior clusters, and important outputs. They should remain
optional configuration layered over the generic capture engine. A semantic map may route review but
must not silently infer user intent or certify business correctness.

## Delivery sequence

1. **0.2 stabilization intake:** collect installer, visual, workflow, and missing-behavior feedback
   from normal maintainer use.
2. **One stabilization pass:** repair the installer/process-handoff issue and accepted 0.2 intake
   findings together; verify the installed candidate without touching private workbooks or reports.
3. **Scenario-recording specification:** define the local scenario record, start/finish state machine,
   lever/output selection, and evidence review.
4. **Desktop scenario library:** implement named scenario persistence and comparison in the host.
5. **Thin Excel add-in:** add in-workbook controls against the proven host contract.
6. **Downstream integration:** define a versioned export consumed by visualization products.

Each phase requires its own approval. Later phases do not expand the scope of the current local
candidate.

## Deferred and unresolved

- Exact add-in technology and local-host communication contract.
- Whether scenario names bound to cells are captured as values, references, or both.
- Behavior when selected ranges move, sheets are renamed, or a workbook is saved under a new name.
- Export schema and trust boundary for downstream visualization.
- Public-release qualification, signing, distribution, and support model.

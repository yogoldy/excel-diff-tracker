# Scenario product discovery — 2026-09-03

## Purpose

This note preserves the reasoning from a high-value product brainstorm. It is a synthesis rather
than a verbatim transcript. Examples are generalized so the public repository does not contain
institutional workbooks, values, communications, or other private evidence.

## Starting observation

The discussion began while examining how multiple decision scenarios can share an established Excel
model. A model may contain a stable bridge into downstream calculations while scenario-specific
inputs, assumptions, explanations, and controls live on separate user-facing sheets.

That exposed several recurring risks:

- one sheet can unintentionally serve as user interface, calculation engine, override store, router,
  and results page at the same time;
- saved inputs for one scenario can remain live while another scenario is selected;
- a scenario can borrow an important assumption from an unrelated sheet without displaying it;
- one generic bridge cell can acquire different meanings in different scenarios;
- a mode labeled “baseline” may disable only one overlay rather than every scenario effect; and
- leadership explanations can omit values that still affect the authoritative model.

The architectural response was to preserve stable downstream bridge addresses while separating
scenario ownership and adding one central gate, rather than adding switch logic throughout every
downstream sheet.

## Product leap: from one model to a general tool

The immediate question was whether workbook-diff technology could supplement analyst-led scenario
development and become useful to other analysts. The answer was refined from “AI that understands a
spreadsheet” to a more controlled proposition:

> Use the existing workbook as the calculation authority; add baseline discipline, scenario
> execution, difference evidence, validation, and decision communication around it.

Three kinds of difference matter:

1. **Structural difference:** formulas, values, sheets, names, VBA, or workbook structure changed.
2. **Runtime difference:** outputs changed after approved inputs were applied and Excel recalculated.
3. **Decision or semantic difference:** the business mechanisms, timing, assumptions, evidence, and
   limitations that explain material runtime changes.

Raw cell comparison is useful but insufficient. The differentiated product would convert a large
mechanical diff into a small, analyst-approved account of what changed, why, when, and whether it
reconciles.

## Trust discussion

AI was treated as supplementary rather than authoritative. A sophisticated AI agent may help an
analyst develop or inspect a model, but large organizations are unlikely to let management rely on an
agent to reinterpret sensitive, complex workbook logic on every run.

The accepted direction was therefore:

- analyst-supervised AI may assist configuration and explanation;
- the analyst approves the scenario contract;
- the runtime writes only approved values;
- desktop Excel performs calculation;
- fixed validations and reconciliations determine run status; and
- local-only execution addresses privacy and institutional-data concerns.

This produced the read/operate/build distinction documented in
[Trust and execution boundary](../TRUST_AND_EXECUTION_BOUNDARY.md).

## Market and competition discussion

Initial competitor research covered workbook comparison, model audit, spreadsheet governance, AI in
Excel, and enterprise FP&A. The list was then narrowed according to the actual product definition.

- Microsoft comparison tools are useful features but do not manage a decision on one workbook.
- Copilot and other agents are development aids, not the trusted management runtime.
- PerfectXL and Operis OAK merit hands-on review and may offer valuable product lessons.
- Cube and Datarails are valid strategic concerns because they preserve Excel while adding planning
  and governance.
- Vena is a broader Excel-oriented platform alternative.
- ClusterSeven is governance rather than scenario operation.
- Workday Adaptive Planning and Anaplan make a different planning engine authoritative.

The discussion concluded that there is a plausible market-shaped gap but not yet proof of a market.
Adjacent spending demonstrates pain; only user behavior and willingness to pay can validate this
specific product.

## The onboarding problem and latent scenario models

A major concern was that every customer's workbook might require bespoke semantic consulting. The
key counter-observation was that many ordinary forecasts already contain:

```text
inputs → formulas → outputs
```

They may lack scenario sheets, switches, saved cases, or formal governance, but analysts already
change assumptions and inspect the consequences. Excel is already their informal scenario engine.

The software therefore does not initially need to understand the complete workbook. The analyst can
identify a few controls and outputs. The product can execute and record that relationship outside the
workbook, while whole-workbook diff remains a safety net for unexpected effects.

This greatly reduces onboarding and expands the addressable hypothesis from organizations with
sophisticated scenario engines to organizations with latent driver-based models and manual what-if
processes.

## Product layers discovered

The discussion converged on a progression:

1. Record a manually performed scenario and export its diff.
2. Replay analyst-approved inputs through Excel and capture approved outputs.
3. Add semantic labels, timing, evidence, leakage checks, and reconciliation.
4. Add team governance and explicit portfolio combinations.

These layers are detailed in [Product layers](../PRODUCT_LAYERS.md).

## The simple multi-sheet scenario-manager wedge

An additional product possibility emerged: the built-in Excel scenario-management experience is a
poor fit for analysts who need to organize inputs and outputs distributed across many worksheets. A
small local add-in could offer:

- click-to-select inputs and outputs across sheets;
- readable labels and units;
- baseline capture and one-click restoration;
- named, duplicable scenarios;
- side-by-side result comparison; and
- export into GCBS visualization.

This may support a straightforward search-led, low-cost product before the full corporate assurance
system exists. It also creates useful scenario records for the larger product, so it is not throwaway
work.

## Important product-design distinctions retained

- **Value writes are not logic writes.** Replaying approved values is materially different from
  allowing software or AI to invent formulas and business relationships.
- **Recording works more broadly than replay.** A static or weakly modeled workbook can still have
  manual changes recorded, but it cannot support trustworthy automated scenarios without existing
  driver logic.
- **A whole-workbook diff remains the safety net.** Selected outputs organize attention but must not
  become the only detection boundary.
- **The workbook remains independently usable.** Scenario infrastructure should supplement rather
  than trap the model.
- **GCBS should receive a structured scenario package.** It should not need to interpret an entire
  sensitive workbook.
- **The analyst is the semantic authority.** Software evidence can corroborate intent and mechanics;
  it cannot establish business correctness by itself.
- **Management plug-and-play happens after analyst configuration.** Ease of operation must not be
  confused with automatic model understanding.

## Decisions and accepted directions

- Preserve the local-first privacy posture.
- Keep mechanical evidence, user intent, and business correctness distinct.
- Preserve whole-workbook diff even when users select scenario levers and outputs.
- Treat AI as optional analyst assistance, not the calculation authority.
- Explore an analyst-configured, deterministic scenario runner that uses existing workbook logic.
- Explore a simple multi-sheet Excel scenario manager as an entry product.
- Treat GCBS as a downstream consumer of deliberate scenario exports.
- Review PerfectXL and Operis OAK closely; treat Cube and Datarails as strategic adjacent products.

These are product directions and research priorities, not claims of implemented behavior.

## Open questions

- Is scenario execution performed through an Excel add-in, desktop automation, or a coordinated
  combination of both?
- What is the minimum safe scenario contract?
- How is a contract preserved when sheets are renamed, cells move, or formulas change?
- Should the first paid product be a low-cost utility, a professional analyst tool, a
  service-assisted offering, or a combination?
- Can a low one-time price sustain signing, distribution, compatibility work, and support?
- Which output-discovery features can remain deterministic, and where is optional AI useful?
- How should temporary working copies, restoration, macros, external links, and data refresh be
  governed?
- What exact schema should GCBS consume?
- How many target analysts run material what-if analyses frequently enough to pay?
- How much configuration remains customer-specific after they identify their inputs and outputs?

## Next research, separate from current implementation

1. Use the installed 0.2 candidate normally and complete its already-authorized stabilization
   intake before expanding implementation scope.
2. Conduct hands-on research of PerfectXL and OAK without assuming their public positioning captures
   every capability.
3. Interview analysts about their existing multi-sheet what-if workflow, recurrence, review burden,
   privacy constraints, and failures.
4. Test search intent and willingness to install a local multi-sheet scenario utility.
5. Specify the scenario record and trust boundary before implementing workbook writes or Excel
   automation.

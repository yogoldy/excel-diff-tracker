# Product thesis

## The problem

Many consequential Excel workbooks already contain usable scenario logic even when they were not
designed as formal scenario systems. An analyst changes a few assumptions, Excel propagates those
changes through existing formulas, and the analyst manually discovers and communicates the result.

The missing layer is often not calculation. It is controlled execution, provenance, comparison,
validation, and communication.

## Core thesis

> Many Excel workbooks are latent scenario models. Excel already supplies the calculation engine;
> the product can supply a scenario-management and decision-explanation layer without rebuilding the
> workbook.

The analyst supplies a small amount of meaning they already know:

- which cells are legitimate decision levers;
- which cells or ranges are important outputs;
- which saved state is the baseline; and
- what labels, units, and constraints make those controls understandable.

The workbook supplies its existing business logic. Desktop Excel may later supply authoritative
recalculation. The application supplies repeatability, evidence, comparison, safeguards, and export.

## Two related product wedges

### Simple Excel utility

A low-friction, local multi-sheet scenario manager could let an analyst select inputs and outputs
across a workbook, save named what-if cases, restore the baseline, compare results, and export them.
Its promise is intentionally simple:

> Select the cells. Save the scenario. Let Excel calculate. Compare and present the result.

This is a search-discoverable utility and potential low-cost entry product. It improves an existing
Excel workflow without requiring AI, server hosting, workbook surgery, or an enterprise planning
implementation.

### Corporate scenario assurance

The deeper product adds analyst certification, scenario ownership, evidence classifications,
recurring and one-time timing, leakage controls, reconciliation, shared libraries, approvals, and
leadership-versus-audit views. Management operates constrained scenarios that an analyst has already
configured and approved.

The simple utility and corporate product are not competing ideas. The utility creates the scenario
definitions and execution evidence that the assurance product can govern.

## Role of AI

The product is not an AI calculation engine. AI may help an analyst inspect a workbook, suggest
candidate outputs, draft labels, explain formula paths, propose tests, or prepare narrative. The
analyst approves the semantic map and the workbook remains responsible for the numerical result.

The trusted runtime should remain useful with AI disabled.

## Role of downstream visualization

GCBS or another presentation tool should consume a deliberate scenario package rather than interpret
an entire workbook. That package can contain the baseline, scenario inputs, observed outputs,
differences, time series, units, evidence status, validation status, and approved explanatory notes.

The intended division is:

```text
Excel workbook       Existing calculation logic
Scenario tool        Controlled execution, evidence, comparison, and validation
GCBS                  Decision visualization and communication
```

## Market hypothesis

The initial customer is not every Excel user. It is an analyst responsible for an important,
long-lived workbook who repeatedly answers management what-if questions, must explain the result to
other people, and cannot justify replacing the workbook with a broad planning platform.

The critical market question is:

> How many consequential corporate Excel workbooks contain driver-based logic but still rely on a
> manual, poorly recorded what-if process?

That proposition requires direct customer validation; adjacent software spending alone does not
prove a standalone market.

# QuantitativeEngine: Legacy VBA to C# / Excel-DNA Migration

## Project Overview

**QuantitativeEngine** is a portfolio project demonstrating the controlled migration of legacy Excel/VBA quantitative-finance models into a compiled C#/.NET calculation library exposed back to Microsoft Excel through Excel-DNA `.xll` functions.

The objective is not simply to translate VBA syntax into C#. The migration treats each legacy workbook as a specification and regression benchmark, then refactors the quantitative logic into reusable, testable components.

The project currently includes:

- Black-Scholes European option pricing
- Black-Scholes option Greeks: Delta, Gamma, Vega, Theta and Rho
- Excel-DNA worksheet functions for direct use from Excel
- Regression testing against outputs from the original VBA workbooks
- A foundation for implied volatility, numerical pricing, Monte Carlo simulation and portfolio risk analytics

---

## Why Migrate a Legacy VBA Quantitative Model?

VBA remains useful for spreadsheet automation and rapid prototyping, but a quantitative calculation library benefits from moving core pricing and risk logic out of individual workbooks.

The migration aims to improve the following areas.

### 1. Separation of calculation logic from the workbook

In the legacy design, model logic is embedded inside Excel/VBA. In the migrated design, Excel becomes a user interface while the pricing and risk calculations live in a C# library.

```text
Legacy
Excel Workbook
     |
     +-- VBA pricing logic
     +-- VBA Greeks logic
     +-- worksheet inputs/outputs

Migrated
Excel Workbook
     |
     v
Excel-DNA .xll
     |
     v
C# QuantitativeEngine
     +-- Black-Scholes pricing
     +-- Greeks
     +-- future analytics
```

This makes the quantitative engine reusable outside a particular workbook.

### 2. Strong typing and clearer interfaces

C# provides strongly typed method signatures, classes and enumerations. Inputs and outputs can therefore be expressed explicitly rather than depending on loosely structured spreadsheet/VBA state.

For example, the engine can clearly distinguish concepts such as option type, spot, strike, maturity, interest rate and volatility.

### 3. Reuse instead of duplicated calculations

A literal VBA-to-C# translation would preserve duplicated calculations. That is deliberately avoided.

The migrated Greeks module reuses the existing Black-Scholes normal-distribution implementation, while common quantities such as `d1`, `d2` and the normal probability density are calculated through shared helper methods.

The goal is:

> **Preserve the mathematics, improve the software architecture.**

### 4. Source control and auditability

The C# source is maintained in Git and GitHub. Each migration can therefore be committed independently with a meaningful history showing how the quantitative library evolved.

This provides a clearer audit trail than maintaining multiple independent workbook copies.

### 5. Testing and regression control

A migration is not considered successful merely because the C# code compiles.

The original workbook is retained as a **legacy regression benchmark**. Known inputs are run through both implementations and the resulting prices and sensitivities are compared to an appropriate numerical tolerance.

This is particularly important in financial software, where a syntactically correct implementation can still contain a subtle modelling, scaling or convention error.

---

## Migration Methodology

Each legacy model follows the same migration workflow.

### Step 1: Inspect the legacy workbook and VBA

Before coding, identify:

- model inputs
- model outputs
- mathematical formulas
- units and scaling conventions
- call/put behaviour
- numerical approximations
- edge cases
- duplicated functions

The existing model is treated as evidence of the intended legacy behaviour rather than blindly copied line by line.

### Step 2: Establish a golden regression case

A representative set of inputs and the corresponding legacy outputs are recorded before refactoring.

For the migrated Greeks workbook, the benchmark inputs were:

```text
Spot          = 100
Strike        = 95
Time          = 0.25 years
Risk-free rate= 0.062
Volatility    = 0.15
```

The legacy workbook produced approximately:

```text
European Call =  7.20198549
European Put  =  0.74083850

Call Delta    =  0.82331626
Put Delta     = -0.17668374
Gamma         =  0.03457912
Vega          =  0.12967171
Call Theta    = -0.02341970
Put Theta     = -0.00753090
Call Rho      =  0.18782410
Put Rho       = -0.04602303
```

These values provide a golden reference for the migrated implementation.

### Step 3: Preserve financial conventions explicitly

Regression analysis of the original Greeks implementation established important reporting conventions:

- Vega is reported per **1 percentage-point** volatility move.
- Rho is reported per **1 percentage-point** interest-rate move.
- Theta is reported on a **daily** basis using a 365-day divisor.
- Call and put Gamma are identical for the same Black-Scholes inputs.
- Call Delta minus Put Delta equals 1 for this non-dividend-paying Black-Scholes setup.

Preserving conventions such as these is as important as preserving the underlying formula.

### Step 4: Refactor into reusable C# components

The migration currently separates functionality into files such as:

```text
QuantitativeEngine/
    BlackScholesEngine.cs
    OptionGreeks.cs
```

`BlackScholesEngine.cs` provides the core analytical pricing capability and cumulative normal distribution implementation.

`OptionGreeks.cs` builds on that core to provide Delta, Gamma, Vega, daily Theta and Rho.

This architecture avoids recreating an independent Black-Scholes implementation for every new calculation.

### Step 5: Expose selected functions through Excel-DNA

Excel-DNA provides the bridge between the C# calculation engine and Excel.

The migrated functions can be called directly from worksheet cells, for example:

```excel
=QE_Delta(100,95,0.25,0.062,0.15,"Call")
=QE_Gamma(100,95,0.25,0.062,0.15)
=QE_Vega(100,95,0.25,0.062,0.15)
=QE_Theta(100,95,0.25,0.062,0.15,"Call")
=QE_Rho(100,95,0.25,0.062,0.15,"Call")
```

This approach deliberately retains Excel as a familiar front end while moving the calculation logic into compiled C#.

### Step 6: Build and run the `.xll`

The project is compiled and loaded into Excel through Excel-DNA. Successful compilation is only the first checkpoint. The worksheet functions must also be callable and return sensible results.

### Step 7: Regression-test C# against the legacy model

The C# implementation is tested using the same inputs as the original workbook.

Observed C# results for the Greeks migration were:

```text
Call Delta  =  0.823316271
Put Delta   = -0.176683729
Gamma       =  0.03457912
Vega        =  0.129671699
Call Theta  = -0.023419695
Put Theta   = -0.007530904
Call Rho    =  0.187824103
Put Rho     = -0.04602303
```

The results reproduce the legacy outputs to the expected numerical precision.

A useful invariant was also tested:

```text
Call Delta - Put Delta = 1
```

This produced `1` in the Excel-DNA implementation.

Small differences in the final decimal places are expected where implementations use different numerical approximations to functions such as the cumulative normal distribution. Regression tests should therefore use a documented numerical tolerance rather than require arbitrary bit-for-bit equality.

### Step 8: Commit only after successful regression testing

The intended development lifecycle is:

```text
Inspect legacy model
        |
        v
Record benchmark results
        |
        v
Refactor into C#
        |
        v
Build Excel-DNA .xll
        |
        v
Run worksheet regression tests
        |
        v
Compare with legacy outputs
        |
        v
Commit to Git only when tests pass
```

This makes each Git commit a meaningful, tested increment to the engine.

---

## Testing Philosophy

The project uses several levels of verification.

### Legacy parity tests

Does the migrated C# calculation reproduce the known VBA result within tolerance?

### Financial invariants

Tests should also check relationships implied by the model rather than only copying expected numbers.

Examples include:

```text
Call Delta - Put Delta = 1
Call Gamma = Put Gamma
Call Vega  = Put Vega
```

### Behavioural tests

Future tests should vary:

- deep in-the-money and out-of-the-money options
- short and long maturities
- low and high volatility
- low and high rates
- call and put option types

### Edge-case and input validation

The engine should progressively add explicit handling for invalid or limiting inputs such as:

- non-positive spot
- non-positive strike
- negative maturity
- zero or negative volatility where unsupported
- invalid option-type strings

### Automated unit testing

The current Excel regression sheet is a valuable migration test harness. As the engine grows, the same golden cases should also be captured in an automated C# test project so that future changes can be regression-tested without manually opening Excel.

---

## Why Excel-DNA?

The purpose of the migration is **not to remove Excel**.

Excel remains extremely useful as an interactive analysis and reporting environment. Excel-DNA allows the project to preserve that interface while replacing embedded VBA calculations with C#/.NET functions exposed through an `.xll` add-in.

Conceptually:

```text
Excel = presentation, interaction and ad-hoc analysis

Excel-DNA = integration boundary

C#/.NET = quantitative calculation engine
```

This separation is valuable because the calculation library can evolve independently from individual workbooks.

---

## Benefits Over the Legacy Architecture

### Legacy workbook-centric approach

- calculation code distributed through VBA modules/workbooks
- greater coupling between model and spreadsheet
- repeated implementation of common mathematics
- manual comparison between workbook versions
- limited separation between calculation, interface and orchestration

### Migrated engine approach

- reusable C# calculation library
- strongly typed interfaces
- common mathematical functions shared across models
- Git-based source history
- regression tests against legacy outputs
- Excel retained through Excel-DNA
- architecture suitable for extension beyond a single workbook
- a pathway toward automated unit and integration tests

The project does **not** assume that C# automatically makes a model financially correct. Correctness comes from understanding the original model, preserving required conventions, testing against known results and reviewing differences deliberately.

---

## Current Architecture

```text
                     Microsoft Excel
                           |
                           | worksheet UDF calls
                           v
                    Excel-DNA (.xll)
                           |
                           v
                QuantitativeEngine (.NET)
                    /              \
                   /                \
      BlackScholesModel          OptionGreeks
             |                  / / | \\ \
             |             Delta Gamma Vega Theta Rho
             |
             +---- shared normal CDF / pricing logic
```

The next stage will extend this architecture rather than create isolated functions.

---

## Planned Development Roadmap

A logical progression for the original VBA collection is:

### Completed

- [x] Black-Scholes European option pricing
- [x] Option Greeks
- [x] Excel-DNA integration
- [x] Regression comparison with legacy VBA outputs

### Next

- [ ] Implied volatility using Newton-Raphson
- [ ] Alternative implied-volatility solvers such as bisection and secant methods
- [ ] Binomial-tree pricing
- [ ] Random-number and stochastic-process components
- [ ] Monte Carlo option pricing
- [ ] Variance-reduction techniques
- [ ] Portfolio-level pricing and aggregated Greeks
- [ ] Scenario and stress calculations
- [ ] Automated C# unit/regression test suite
- [ ] Volatility smile and surface components

A particularly important architectural objective for implied volatility is reuse. Newton-Raphson can consume both the existing Black-Scholes pricing function and Vega calculation rather than introduce duplicate implementations.

---

## Migration Principles

The project follows several simple rules.

1. **Understand before translating.**
2. **Treat the legacy workbook as a regression specification, not as ideal software architecture.**
3. **Preserve financial meaning, units and conventions.**
4. **Refactor duplicated mathematics into shared components.**
5. **Keep Excel as an interface where it remains useful.**
6. **Move core analytics into a reusable C# library.**
7. **Test numerical outputs before committing migrated code.**
8. **Use financial identities and invariants in addition to point-value regression tests.**
9. **Document deliberate differences from the legacy implementation.**
10. **Modernise incrementally so that every stage remains verifiable.**

---

## Portfolio Perspective

This repository is intended to demonstrate more than knowledge of an option-pricing equation. It documents the process of taking legacy spreadsheet-based quantitative logic and progressively turning it into a structured software-engineering project.

The work combines:

- financial mathematics
- legacy-code analysis
- C#/.NET development
- numerical methods
- Excel interoperability
- regression testing
- Git/GitHub source control
- refactoring and software architecture

The original VBA models provide useful quantitative examples and known outputs. The C# migration provides an opportunity to retain the financial knowledge while applying modern engineering practices around the calculations.

---

## Disclaimer

This project is for educational, portfolio and software-engineering demonstration purposes. It is not investment advice and the models should not be relied upon for production trading or risk management without appropriate independent model validation, testing, controls and governance.

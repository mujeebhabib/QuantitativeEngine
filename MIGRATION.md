# QuantitativeEngine: Legacy VBA to C# / Excel-DNA Migration

## Project Overview

**QuantitativeEngine** is a portfolio project demonstrating the controlled migration of legacy Excel/VBA quantitative-finance models into a compiled C#/.NET calculation library exposed back to Microsoft Excel through Excel-DNA `.xll` functions.

The objective is not simply to translate VBA syntax into C#. Each legacy workbook is treated as a combination of:

- quantitative specification
- historical implementation
- regression benchmark
- source of financial conventions and assumptions

The underlying mathematics is preserved, while the implementation is progressively refactored into reusable, strongly typed and testable C# components.

The project currently includes:

- Black-Scholes European option pricing
- Black-Scholes option Greeks: Delta, Gamma, Vega, Theta and Rho
- Newton-Raphson implied-volatility calibration for European calls
- Excel-DNA worksheet functions exposed through an `.xll` add-in
- regression testing against original VBA workbook outputs
- repricing, financial-invariant and round-trip tests
- Git/GitHub source control and incremental migration history

The guiding principle is:

> **Preserve the mathematics, improve the software architecture.**

---

## Why Migrate a Legacy VBA Quantitative Model?

VBA remains useful for spreadsheet automation and rapid prototyping. The purpose of this project is therefore not to remove Excel or dismiss the original VBA implementations.

Instead, the aim is to separate reusable quantitative calculations from individual workbooks while retaining Excel as a practical analytical interface.

### 1. Separation of calculation logic from the workbook

In the legacy design, quantitative logic and spreadsheet presentation are closely coupled:

```text
Legacy

Excel Workbook
     |
     +-- VBA pricing logic
     +-- VBA Greeks logic
     +-- numerical routines
     +-- worksheet inputs/outputs
```

The migrated architecture separates these responsibilities:

```text
Migrated

Excel Workbook
     |
     v
Excel-DNA .xll
     |
     v
C# QuantitativeEngine
     |
     +-- Black-Scholes pricing
     +-- Option Greeks
     +-- Implied-volatility calibration
     +-- future numerical engines
```

Excel remains the interactive front end, but the financial calculations live in a compiled .NET library.

### 2. Strong typing and clearer interfaces

C# provides strongly typed method signatures, classes and enumerations. Inputs such as spot, strike, maturity, interest rate, volatility and option type can therefore be represented explicitly rather than being dependent on loosely structured workbook state.

### 3. Reuse rather than duplicated calculations

A literal VBA-to-C# translation would preserve duplicated implementations. This project deliberately avoids that approach.

For example:

- the Greeks layer builds on the Black-Scholes mathematical core
- the implied-volatility solver calls the existing Black-Scholes pricing function
- shared quantities such as `d1`, normal-distribution calculations and Vega are treated as reusable quantitative concepts

The target is an engine whose components cooperate rather than a collection of unrelated worksheet functions.

### 4. Source control and auditability

The C# source is maintained in Git and GitHub. Each migration is committed as an incremental change after regression testing.

This creates a history showing how the engine evolved from individual legacy models into a reusable quantitative library.

### 5. Testing and regression control

Successful compilation is not treated as proof of numerical correctness.

The original workbooks provide known inputs and outputs. These are retained as golden regression benchmarks, then supplemented with financial identities, repricing checks and independent round-trip tests.

This is particularly important in financial software because code can compile successfully while still containing a scaling, convention, numerical or modelling error.

---

# Migration Methodology

Each legacy model follows the same controlled migration process.

## Step 1: Inspect the legacy workbook and VBA

Before writing C#, identify:

- model inputs
- model outputs
- mathematical formulas
- reporting units
- scaling conventions
- call/put behaviour
- numerical approximations
- convergence rules where applicable
- edge cases
- duplicated calculations

The legacy model is treated as evidence of intended behaviour, not automatically as the ideal target software architecture.

## Step 2: Establish golden regression benchmarks

Representative legacy inputs and outputs are recorded before refactoring.

This allows the migrated implementation to be evaluated against known behaviour rather than judged by appearance or compilation alone.

## Step 3: Preserve financial meaning and conventions

The migration explicitly identifies conventions that may not be obvious from a formula alone.

Examples discovered during the Greeks migration include:

- Vega reported per **1 percentage-point** volatility move
- Rho reported per **1 percentage-point** interest-rate move
- Theta reported on a **daily** basis using a 365-day divisor
- identical call and put Gamma under the model assumptions
- `Call Delta - Put Delta = 1` for the non-dividend-paying Black-Scholes setup used in the benchmark

These conventions are documented rather than silently changed.

## Step 4: Refactor into reusable C# components

The current quantitative source structure includes:

```text
QuantitativeEngine/
    BlackScholesEngine.cs
    OptionGreeks.cs
    ImpliedVolatility.cs
```

The intention is for each new component to reuse the existing engine wherever appropriate.

## Step 5: Expose selected calculations through Excel-DNA

Excel-DNA provides the integration boundary between the C# calculation engine and Microsoft Excel.

Examples include:

```excel
=PriceOptionUDF(100,95,0.25,0.062,0.15,"Call")
=QE_Delta(100,95,0.25,0.062,0.15,"Call")
=QE_Gamma(100,95,0.25,0.062,0.15)
=QE_Vega(100,95,0.25,0.062,0.15)
=QE_Theta(100,95,0.25,0.062,0.15,"Call")
=QE_Rho(100,95,0.25,0.062,0.15,"Call")
=QE_ImpliedVolatilityCall(100,95,0.25,0.07,10)
```

Excel therefore remains useful for interactive analysis and regression testing without containing the core implementation itself.

## Step 6: Build and run the `.xll`

The C# project is compiled and loaded into Excel through Excel-DNA. Successful compilation is only the first checkpoint. The worksheet functions must execute correctly and produce financially sensible results.

## Step 7: Regression-test against the legacy implementation

The migrated calculation is compared with known outputs from the original workbook.

Small differences caused by numerical approximation are evaluated using appropriate tolerances rather than arbitrary bit-for-bit equality.

## Step 8: Add independent mathematical tests

Legacy parity alone is not sufficient. Where possible, the project also adds tests based on model identities, monotonic behaviour, repricing and known-input round trips.

## Step 9: Commit only after testing

The development lifecycle is:

```text
Inspect legacy model
        |
        v
Record benchmark results
        |
        v
Understand conventions
        |
        v
Refactor into C#
        |
        v
Build Excel-DNA .xll
        |
        v
Run regression tests
        |
        v
Run independent checks
        |
        v
Commit to Git only after successful verification
```

This makes each Git commit a meaningful tested increment to `QuantitativeEngine`.

---

# Migration 1: Black-Scholes European Option Pricing

The first migration established the core analytical pricing engine.

For a European call:

```text
C = S N(d1) - K exp(-rT) N(d2)
```

For a European put:

```text
P = K exp(-rT) N(-d2) - S N(-d1)
```

with:

```text
d1 = [ln(S/K) + (r + sigma^2 / 2)T] / [sigma sqrt(T)]
d2 = d1 - sigma sqrt(T)
```

The migration created `BlackScholesEngine.cs`, providing the pricing functionality and cumulative-normal calculation used by subsequent parts of the engine.

This first stage established the forward pricing relationship:

```text
Spot
Strike
Time
Rate          ---> Black-Scholes ---> Option Price
Volatility
Option Type
```

The importance of this migration extends beyond the first pricing function because later components can call the same engine rather than implement Black-Scholes repeatedly.

---

# Migration 2: Black-Scholes Option Greeks

The second migration added the principal first- and second-order Black-Scholes sensitivities:

- Delta
- Gamma
- Vega
- Theta
- Rho

## Legacy benchmark

The legacy workbook used:

```text
Spot           = 100
Strike         = 95
Time           = 0.25 years
Risk-free rate = 0.062
Volatility     = 0.15
```

and produced approximately:

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

## Migrated C# results

The Excel-DNA implementation produced approximately:

```text
Call Delta    =  0.823316271
Put Delta     = -0.176683729
Gamma         =  0.03457912
Vega          =  0.129671699
Call Theta    = -0.023419695
Put Theta     = -0.007530904
Call Rho      =  0.187824103
Put Rho       = -0.04602303
```

The small final-decimal differences are consistent with numerical approximation differences while preserving the intended quantitative behaviour.

## Financial invariant tests

A particularly useful independent test was:

```text
Call Delta - Put Delta = 1
```

The Excel-DNA implementation returned exactly `1` for the benchmark case.

The implementation also preserves:

```text
Call Gamma = Put Gamma
Call Vega  = Put Vega
```

under the model assumptions used here.

## Scaling conventions

The original VBA established conventions that were intentionally preserved:

```text
Vega  = raw Vega / 100
Rho   = raw Rho / 100
Theta = annual Theta / 365
```

This means the Excel-facing Vega and Rho represent sensitivity to a one percentage-point move, while Theta is presented as daily decay.

---

# Migration 3: Implied Volatility Using Newton-Raphson

The third migration introduced `ImpliedVolatility.cs` and represents an important progression in the project.

Black-Scholes solves the forward problem:

```text
Market Parameters + Volatility
              |
              v
        Black-Scholes
              |
              v
         Option Price
```

Implied volatility solves the inverse problem:

```text
Market Parameters + Observed Option Price
                   |
                   v
            Numerical Solver
                   |
                   v
          Implied Volatility
```

Given spot, strike, maturity, interest rate and an observed option price, the solver determines the volatility that causes the Black-Scholes model to reproduce the observed price.

## Root-finding formulation

Define:

```text
f(sigma) = BlackScholesPrice(sigma) - MarketPrice
```

The required implied volatility is the value of `sigma` for which:

```text
f(sigma) = 0
```

Newton-Raphson updates the volatility estimate using:

```text
sigma(new) = sigma(old)
             - [ModelPrice(sigma(old)) - MarketPrice]
               / RawVega(sigma(old))
```

This creates an important connection between the first three migrations:

```text
Black-Scholes Pricing
        |
        v
Option Greeks / Vega
        |
        v
Newton-Raphson Root Finding
        |
        v
Implied Volatility Calibration
```

The project has therefore progressed from pricing, to sensitivity analysis, to calibration.

## Legacy methodology retained

The legacy VBA Newton implementation was studied before the C# solver was written.

The migrated implementation preserves the important concepts:

- Newton-Raphson iteration
- Vega as the derivative used by Newton's method
- an initial volatility estimate based on the legacy methodology
- a numerical convergence tolerance
- a maximum iteration count
- benchmark comparison with the original workbook
- repricing as an independent verification step

The C# implementation also introduces clearer input validation and explicit handling of solver failure conditions.

## Reuse of the existing pricing engine

`ImpliedVolatility.cs` does not contain another independent Black-Scholes pricing implementation.

Instead, Newton-Raphson repeatedly calls the existing `BlackScholesModel.PriceOptionUDF(...)` calculation while changing `sigma` until the model price converges to the observed market price.

Conceptually:

```text
                    +-----------------------+
                    | BlackScholesEngine.cs |
                    +-----------+-----------+
                                ^
                                | model-price evaluations
                                |
                    +-----------+-----------+
Market Price -----> | ImpliedVolatility.cs  |
                    |   Newton-Raphson       |
                    +-----------+-----------+
                                |
                                v
                         Implied sigma
```

This is an important architectural improvement over duplicating pricing formulas inside each legacy migration.

## Raw Vega versus reported Vega

The Greeks worksheet function reports Vega per one percentage-point move and therefore scales Vega by `/100`.

Newton-Raphson requires the mathematical derivative with respect to raw volatility:

```text
dV / dSigma
```

The solver therefore uses **unscaled raw Vega internally**.

This distinction ensures that an Excel reporting convention does not alter the numerical mathematics of the solver.

---

## Implied-Volatility Legacy Regression Tests

The original Newton-Raphson workbook supplied three benchmark cases.

The common parameters were:

```text
Spot           = 100
Strike         = 95
Time           = 0.25 years
Risk-free rate = 0.07
```

Only the market call price changed.

### Legacy results

```text
Market Price     Legacy Implied Volatility
------------------------------------------
10               0.3172519803
20               0.8606126904
50               2.5870265961
```

### C# results

The new Excel-DNA solver produced:

```text
Market Price     C# Implied Volatility
------------------------------------------
10               0.317254926854488
20               0.860630552427070
50               2.587064732459632
```

The small differences from the legacy implied volatilities are acceptable because the old and new pricing engines use different numerical approximations. The more important question is whether the new volatility solves the pricing equation inside the migrated C# engine.

## Repricing regression

Each C# implied volatility was fed back into `BlackScholesModel` using the same spot, strike, maturity and rate.

The results were:

```text
Observed Market Price     C# Repriced Value
------------------------------------------------
10                        10.000000000146784
20                        19.999999969142735
50                        50.000000000000156
```

The migrated solver therefore satisfies, to very high numerical accuracy:

```text
BlackScholesPrice(ImpliedVolatility) ~= MarketPrice
```

This repricing test is considered stronger than forcing exact equality with every decimal generated by a different legacy numerical implementation.

## Behavioural test

The benchmark cases also provide an intuitive financial test.

Holding spot, strike, maturity and rate constant:

```text
Higher option market price
          |
          v
Higher implied volatility
```

The migrated solver produces:

```text
Market Price 10  -> IV approximately 31.7%
Market Price 20  -> IV approximately 86.1%
Market Price 50  -> IV approximately 258.7%
```

The expected monotonic behaviour is preserved.

## Independent round-trip test

A separate test was performed without relying on a legacy implied-volatility output.

A deliberately known volatility of 20% was supplied to the C# Black-Scholes engine:

```text
Spot           = 100
Strike         = 100
Time           = 1 year
Risk-free rate = 0.05
Volatility     = 0.20
```

The Black-Scholes implementation generated a call value of approximately:

```text
10.45057542
```

That generated price was then treated as a synthetic market price and sent to the Newton-Raphson solver without supplying the original volatility.

The process was:

```text
Known volatility = 20%
          |
          v
C# Black-Scholes
          |
          v
Synthetic option price
          |
          v
C# Newton-Raphson
          |
          v
Recovered volatility ~= 20%
```

The solver successfully recovered approximately 20%.

This provides an independent test of internal consistency between the pricing and calibration components.

---

# Testing Philosophy

The project deliberately uses more than one type of test.

## 1. Legacy parity tests

Does the migrated C# calculation reproduce the behaviour of the original VBA model within an appropriate numerical tolerance?

## 2. Repricing tests

Where a solver produces a model parameter, does feeding that parameter back into the pricing engine recover the target market value?

For implied volatility:

```text
Market Price
     |
     v
Implied Volatility Solver
     |
     v
Implied sigma
     |
     v
Black-Scholes Pricing Engine
     |
     v
Repriced value ~= Market Price
```

## 3. Financial invariants

Tests also check relationships implied by the model rather than relying only on copied benchmark numbers.

Examples include:

```text
Call Delta - Put Delta = 1
Call Gamma = Put Gamma
Call Vega  = Put Vega
```

## 4. Behavioural tests

Financially sensible directional relationships are also tested.

For example, with other Black-Scholes parameters held constant:

```text
Option Price increases -> Implied Volatility increases
```

Future behavioural tests should cover:

- deep in-the-money options
- deep out-of-the-money options
- short and long maturities
- low and high volatility
- low and high rates
- calls and puts

## 5. Round-trip tests

Where possible, a known parameter is used to generate a model output and then recovered through the inverse calculation.

The 20% implied-volatility round trip is the first example of this approach.

## 6. Edge-case and input validation

The engine progressively adds explicit handling for invalid or numerically difficult inputs, including:

- non-positive spot
- non-positive strike
- non-positive maturity where unsupported
- zero or negative volatility where unsupported
- invalid option types
- impossible or inconsistent market prices
- near-zero Vega during Newton iteration
- failure to converge within the maximum iteration count
- `NaN` or infinite numerical results

## 7. Automated unit testing

The Excel regression sheets currently provide a useful migration test harness because legacy and migrated calculations can be inspected together.

As the engine grows, the same golden cases should also be moved into an automated C# test project so future changes can run the regression suite without requiring manual Excel testing.

---

# Why Excel-DNA?

The migration is **not intended to remove Excel**.

Excel remains valuable as an interactive analytical, reporting and scenario-testing environment. Excel-DNA allows the project to retain that familiar interface while moving the reusable quantitative logic into compiled C#/.NET code exposed through an `.xll` add-in.

Conceptually:

```text
Excel
  = presentation, interaction and ad-hoc analysis

Excel-DNA
  = integration boundary

C#/.NET QuantitativeEngine
  = pricing, risk, numerical methods and calibration
```

This separation allows the quantitative library to evolve independently from individual spreadsheets.

---

# Benefits Over the Legacy Architecture

## Legacy workbook-centric approach

- quantitative code distributed through VBA modules and workbooks
- close coupling between model logic and spreadsheet presentation
- repeated implementations of common mathematics
- workbook-specific functions
- manual workbook version comparison
- limited separation between calculation, interface and orchestration

## Migrated engine approach

- reusable C# quantitative library
- strongly typed interfaces
- shared mathematical components
- compiled Excel-DNA integration
- Git-based source history
- regression testing against legacy outputs
- independent mathematical and behavioural tests
- clearer separation of calculation and presentation
- a pathway to automated unit and integration testing
- architecture suitable for additional pricing and risk engines

The project does **not** assume that C# automatically makes a financial model correct.

Correctness comes from understanding the model, preserving required financial conventions, testing against known behaviour and explicitly reviewing differences.

---

# Current Architecture

The engine has now progressed beyond isolated functions:

```text
                         Microsoft Excel
                               |
                               | worksheet UDF calls
                               v
                        Excel-DNA (.xll)
                               |
                               v
                    QuantitativeEngine (.NET)
                               |
             +-----------------+----------------+
             |                                  |
             v                                  v
   BlackScholesModel                     OptionGreeks
             |                       Delta Gamma Vega
             |                           Theta Rho
             |
             +-----------------+
                               |
                               v
                      ImpliedVolatility
                        Newton-Raphson
                               |
                               v
                     Calibrated implied sigma
```

The important architectural progression is:

```text
Pricing
   |
   v
Sensitivities
   |
   v
Numerical Root Finding
   |
   v
Calibration
```

Future migrations will continue extending this engine rather than creating independent one-off implementations.

---

# Numerical Types

Core quantitative calculations currently use C# `double` precision.

This is appropriate for numerical operations involving functions such as logarithms, exponentials, square roots, normal distributions and iterative solvers.

The project distinguishes quantitative modelling from exact decimal accounting arithmetic. A future cash, settlement or accounting layer may use `decimal` where appropriate, while model calculations remain based on numerical floating-point methods.

A useful project-level distinction is:

```text
Quantitative modelling     -> double
Cash/accounting amounts    -> consider decimal where appropriate
```

---

# Planned Development Roadmap

A logical progression for the original VBA collection and `QuantitativeEngine` is:

## Completed

- [x] Black-Scholes European option pricing
- [x] Black-Scholes option Greeks
- [x] Newton-Raphson implied volatility for European calls
- [x] Excel-DNA `.xll` integration
- [x] Regression comparison with legacy VBA outputs
- [x] Greeks financial-invariant tests
- [x] Implied-volatility repricing tests
- [x] Implied-volatility independent round-trip test
- [x] Git/GitHub incremental migration history

## Next

- [ ] Automated C# unit/regression test project
- [ ] Alternative implied-volatility solvers using bisection and secant methods
- [ ] Put implied-volatility support
- [ ] Binomial-tree pricing
- [ ] American-option pricing exercises
- [ ] Random-number and stochastic-process components
- [ ] Monte Carlo option pricing
- [ ] Variance-reduction techniques
- [ ] Portfolio-level pricing and aggregated Greeks
- [ ] Scenario and stress calculations
- [ ] Volatility smile analysis
- [ ] Volatility surface components and interpolation
- [ ] Release-build and deployment workflow for the Excel-DNA add-in

The long-term aim is not to migrate every VBA file mechanically. Legacy examples are prioritised where the underlying concept remains useful and where the migration extends the quantitative architecture in a meaningful way.

---

# Migration Principles

The project follows ten core principles:

1. **Understand before translating.**
2. **Treat the legacy workbook as a regression specification, not as ideal software architecture.**
3. **Preserve financial meaning, units and conventions.**
4. **Refactor duplicated mathematics into shared components.**
5. **Keep Excel as an interface where Excel remains useful.**
6. **Move reusable quantitative analytics into the C# engine.**
7. **Test numerical outputs before committing migrated code.**
8. **Use financial identities, repricing and round trips in addition to point-value tests.**
9. **Document deliberate differences from the legacy implementation.**
10. **Modernise incrementally so each stage remains understandable and verifiable.**

---

# Portfolio Perspective

This repository is intended to demonstrate more than knowledge of an option-pricing equation.

It documents the progressive transformation of legacy spreadsheet-based quantitative logic into a structured software-engineering project combining:

- financial mathematics
- derivatives pricing
- risk sensitivities
- numerical root finding
- model calibration
- legacy-code analysis
- C#/.NET development
- Excel interoperability
- Excel-DNA add-in development
- numerical regression testing
- Git/GitHub source control
- refactoring and software architecture

The legacy VBA models provide useful quantitative specifications and benchmark outputs. The migration retains that financial knowledge while applying more maintainable engineering practices around the calculations.

The architecture is intended to evolve progressively:

```text
Legacy VBA Models
       |
       v
Black-Scholes Pricing
       |
       v
Option Greeks
       |
       v
Implied Volatility Calibration
       |
       v
Binomial / Numerical Pricing
       |
       v
Monte Carlo Simulation
       |
       v
Portfolio Risk and Scenario Analytics
```

Each stage should build on and test the components created before it.

---

# Disclaimer

This project is for educational, portfolio and software-engineering demonstration purposes.

The calculations are not investment advice and should not be relied upon for live trading, valuation, hedging or risk management without appropriate independent model validation, testing, market-data controls, governance and production engineering.

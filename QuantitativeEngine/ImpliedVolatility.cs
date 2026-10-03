using ExcelDna.Integration;
using System;

namespace QuantitativeEngine
{
    public static class ImpliedVolatility
    {
        private const double DefaultTolerance = 0.0000001;
        private const int MaxIterations = 100;
        private const double MinimumVega = 1e-12;

        /// <summary>
        /// Standard normal probability density function.
        /// Used to calculate raw Black-Scholes Vega.
        /// </summary>
        private static double NormalPdf(double x)
        {
            return Math.Exp(-0.5 * x * x)
            / Math.Sqrt(2.0 * Math.PI);
        }

        /// <summary>
        /// Calculates d1 for the Black-Scholes model.
        /// </summary>
        private static double D1(
        double spot,
        double strike,
        double time,
        double rate,
        double sigma)
        {
            return (
            Math.Log(spot / strike)
            + (rate + 0.5 * sigma * sigma) * time
            ) / (sigma * Math.Sqrt(time));
        }

        /// <summary>
        /// Raw Black-Scholes Vega: dV/dSigma.
        ///
        /// IMPORTANT:
        /// This is deliberately NOT divided by 100.
        /// Newton-Raphson requires the mathematical derivative
        /// with respect to raw sigma.
        /// </summary>
        private static double RawVega(
        double spot,
        double strike,
        double time,
        double rate,
        double sigma)
        {
            double d1 = D1(
            spot,
            strike,
            time,
            rate,
            sigma);

            return spot
            * NormalPdf(d1)
            * Math.Sqrt(time);
        }

        /// <summary>
        /// Calculates the Brenner-Subrahmanyam-style initial
        /// volatility estimate used by the legacy VBA model.
        /// </summary>
        private static double InitialGuess(
        double spot,
        double time,
        double marketPrice)
        {
            return marketPrice
            / (0.398 * spot * Math.Sqrt(time));
        }

        /// <summary>
        /// Calculates Black-Scholes implied volatility for a
        /// European call using Newton-Raphson iteration.
        /// </summary>
        [ExcelFunction(
        Name = "QE_ImpliedVolatilityCall",
        Description =
        "Calculates Black-Scholes call implied volatility using Newton-Raphson")]
        public static double ImpliedVolatilityCall(
        double spot,
        double strike,
        double time,
        double rate,
        double marketPrice)
        {
            ValidateInputs(
            spot,
            strike,
            time,
            marketPrice);

            double sigma = InitialGuess(
            spot,
            time,
            marketPrice);

            // Defensive fallback in case the initial estimate
            // is numerically unusable.
            if (sigma <= 0.0 ||
            double.IsNaN(sigma) ||
            double.IsInfinity(sigma))
            {
                sigma = 0.20;
            }

            for (int iteration = 1;
            iteration <= MaxIterations;
            iteration++)
            {
                double modelPrice =
                BlackScholesModel.PriceOptionUDF(
                spot,
                strike,
                time,
                rate,
                sigma,
                "Call");

                double priceError =
                modelPrice - marketPrice;

                // Economic convergence test:
                // has the model reproduced the market price?
                if (Math.Abs(priceError) < DefaultTolerance)
                {
                    return sigma;
                }

                double vega = RawVega(
                spot,
                strike,
                time,
                rate,
                sigma);

                if (Math.Abs(vega) < MinimumVega)
                {
                    throw new InvalidOperationException(
                    "Newton-Raphson cannot continue because Vega is too close to zero.");
                }

                double correction =
                priceError / vega;

                double newSigma =
                sigma - correction;

                if (newSigma <= 0.0 ||
                double.IsNaN(newSigma) ||
                double.IsInfinity(newSigma))
                {
                    throw new InvalidOperationException(
                    "Newton-Raphson produced an invalid volatility.");
                }

                // Numerical convergence test:
                // is the latest volatility correction tiny?
                if (Math.Abs(newSigma - sigma)
                < DefaultTolerance)
                {
                    return newSigma;
                }

                sigma = newSigma;
            }

            throw new InvalidOperationException(
            "Implied volatility did not converge within the maximum number of iterations.");
        }

        /// <summary>
        /// Basic validation before starting the solver.
        /// </summary>
        private static void ValidateInputs(
        double spot,
        double strike,
        double time,
        double marketPrice)
        {
            if (spot <= 0.0)
            {
                throw new ArgumentOutOfRangeException(
                nameof(spot),
                "Spot must be greater than zero.");
            }

            if (strike <= 0.0)
            {
                throw new ArgumentOutOfRangeException(
                nameof(strike),
                "Strike must be greater than zero.");
            }

            if (time <= 0.0)
            {
                throw new ArgumentOutOfRangeException(
                nameof(time),
                "Time to expiry must be greater than zero.");
            }

            if (marketPrice <= 0.0)
            {
                throw new ArgumentOutOfRangeException(
                nameof(marketPrice),
                "Market option price must be greater than zero.");
            }
        }
    }
}
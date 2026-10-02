using ExcelDna.Integration;
using System;

namespace QuantitativeEngine
{
    public static class OptionGreeks
    {
        private static double D1(
        double spot,
        double strike,
        double time,
        double rate,
        double sigma)
        {
            return (Math.Log(spot / strike)
            + (rate + 0.5 * sigma * sigma) * time)
            / (sigma * Math.Sqrt(time));
        }

        private static double D2(
        double spot,
        double strike,
        double time,
        double rate,
        double sigma)
        {
            return D1(spot, strike, time, rate, sigma)
            - sigma * Math.Sqrt(time);
        }

        private static double NormalPdf(double x)
        {
            return Math.Exp(-0.5 * x * x)
            / Math.Sqrt(2.0 * Math.PI);
        }

        [ExcelFunction(
        Name = "QE_Delta",
        Description = "Black-Scholes Delta for a European call or put")]
        public static double Delta(
        double spot,
        double strike,
        double time,
        double rate,
        double sigma,
        string optionType)
        {
            double d1 = D1(spot, strike, time, rate, sigma);

            if (optionType.Equals(
            "Call",
            StringComparison.OrdinalIgnoreCase))
            {
                return BlackScholesModel
                .CumulativeNormalDistribution(d1);
            }

            return BlackScholesModel
            .CumulativeNormalDistribution(d1) - 1.0;
        }

        [ExcelFunction(
        Name = "QE_Gamma",
        Description = "Black-Scholes Gamma for a European option")]
        public static double Gamma(
        double spot,
        double strike,
        double time,
        double rate,
        double sigma)
        {
            double d1 = D1(spot, strike, time, rate, sigma);

            return NormalPdf(d1)
            / (spot * sigma * Math.Sqrt(time));
        }

        [ExcelFunction(
        Name = "QE_Vega",
        Description = "Black-Scholes Vega per 1 percentage-point volatility move")]
        public static double Vega(
        double spot,
        double strike,
        double time,
        double rate,
        double sigma)
        {
            double d1 = D1(spot, strike, time, rate, sigma);

            return spot
            * NormalPdf(d1)
            * Math.Sqrt(time)
            / 100.0;
        }

        [ExcelFunction(
        Name = "QE_Theta",
        Description = "Black-Scholes daily Theta for a European call or put")]
        public static double Theta(
        double spot,
        double strike,
        double time,
        double rate,
        double sigma,
        string optionType)
        {
            double d1 = D1(spot, strike, time, rate, sigma);
            double d2 = D2(spot, strike, time, rate, sigma);

            double decay =
            -(spot * NormalPdf(d1) * sigma)
            / (2.0 * Math.Sqrt(time));

            if (optionType.Equals(
            "Call",
            StringComparison.OrdinalIgnoreCase))
            {
                return (
                decay
                - rate * strike
                * Math.Exp(-rate * time)
                * BlackScholesModel
                .CumulativeNormalDistribution(d2)
                ) / 365.0;
            }

            return (
            decay
            + rate * strike
            * Math.Exp(-rate * time)
            * BlackScholesModel
            .CumulativeNormalDistribution(-d2)
            ) / 365.0;
        }

        [ExcelFunction(
        Name = "QE_Rho",
        Description = "Black-Scholes Rho per 1 percentage-point interest-rate move")]
        public static double Rho(
        double spot,
        double strike,
        double time,
        double rate,
        double sigma,
        string optionType)
        {
            double d2 = D2(spot, strike, time, rate, sigma);

            if (optionType.Equals(
            "Call",
            StringComparison.OrdinalIgnoreCase))
            {
                return strike
                * time
                * Math.Exp(-rate * time)
                * BlackScholesModel
                .CumulativeNormalDistribution(d2)
                / 100.0;
            }

            return -strike
            * time
            * Math.Exp(-rate * time)
            * BlackScholesModel
            .CumulativeNormalDistribution(-d2)
            / 100.0;
        }
    }
}
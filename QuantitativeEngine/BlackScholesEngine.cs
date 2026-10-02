using ExcelDna.Integration;
using System;
using System.Collections.Generic;
using System.Linq;

namespace QuantitativeEngine
{
    public enum OptionType { Call, Put }

    // Demonstrates advanced multi-threaded data array processing via LINQ
    public class PortfolioRiskSuite
    {
        [ExcelFunction(Description = "Aggregates total market value of a structured options portfolio via LINQ")]
        public static double AggregateBookValue()
        {
            // Simulating an enterprise database array extract (e.g., from SQL or Credit Suisse legacy feed)
            var portfolio = new List<BookPosition>
            {
                new() { AssetId = "CS-USD-C100", Quantity = 500, Spot = 100, Strike = 100, Time = 1.0, Rate = 0.05, Volatility = 0.20, TypeString = "Call" },
                new() { AssetId = "MZH-GBP-P95",  Quantity = 250, Spot = 98,  Strike = 95,  Time = 0.5, Rate = 0.04, Volatility = 0.25, TypeString = "Put" },
                new() { AssetId = "DB-EUR-C120",  Quantity = 1000, Spot = 115, Strike = 120, Time = 0.25, Rate = 0.03, Volatility = 0.30, TypeString = "Call" }
            };

            // A single declarative line of LINQ code replaces hundreds of lines of slow, nested VBA loops
            return portfolio.Sum(pos => BlackScholesModel.PriceOptionUDF(pos.Spot, pos.Strike, pos.Time, pos.Rate, pos.Volatility, pos.TypeString) * pos.Quantity);
        }
    }

    public class BlackScholesModel
    {
        public static double CumulativeNormalDistribution(double z)
        {
            double p = 0.3275911;
            double a1 = 0.254829592; double a2 = -0.284496736; double a3 = 1.421413741;
            double a4 = -1.453152027; double a5 = 1.061405429;

            int sign = (z < 0) ? -1 : 1;
            double x = Math.Abs(z) / Math.Sqrt(2.0);
            double t = 1.0 / (1.0 + p * x);

            double erf = 1.0 - (((((a5 * t + a4) * t + a3) * t + a2) * t + a1) * t * Math.Exp(-x * x));
            return 0.5 * (1.0 + sign * erf);
        }

        // 1. The Core Calculation Library Excel-DNA Function
        [ExcelFunction(Description = "Calculates the analytical Black-Scholes price for an option contract")]
        public static double PriceOptionUDF(double spot, double strike, double time, double rate, double sigma, string typeString)
        {
            if (time <= 0) return Math.Max(0.0, typeString.Equals("Call", StringComparison.OrdinalIgnoreCase) ? spot - strike : strike - spot);

            OptionType type = typeString.Equals("Call", StringComparison.OrdinalIgnoreCase) ? OptionType.Call : OptionType.Put;

            double d1 = (Math.Log(spot / strike) + (rate + (sigma * sigma) / 2.0) * time) / (sigma * Math.Sqrt(time));
            double d2 = d1 - sigma * Math.Sqrt(time);

            if (type == OptionType.Call)
            {
                return spot * CumulativeNormalDistribution(d1) - strike * Math.Exp(-rate * time) * CumulativeNormalDistribution(d2);
            }
            else
            {
                return strike * Math.Exp(-rate * time) * CumulativeNormalDistribution(-d2) - spot * CumulativeNormalDistribution(-d1);
            }
        }
    }

    // 2. The Portfolio & LINQ Aggregator Layer
    public class BookPosition
    {
        public string AssetId { get; set; } = string.Empty; // Fixed warning
        public double Quantity { get; set; }
        public double Spot { get; set; }
        public double Strike { get; set; }
        public double Time { get; set; }
        public double Rate { get; set; }
        public double Volatility { get; set; }
        public string TypeString { get; set; } = string.Empty; // Fixed warning
    }

    public class RiskManagementSuite
    {
        public double AggregateBookValue(List<BookPosition> portfolio)
        {
            return portfolio.Sum(pos => BlackScholesModel.PriceOptionUDF(pos.Spot, pos.Strike, pos.Time, pos.Rate, pos.Volatility, pos.TypeString) * pos.Quantity);
        }
    }

}
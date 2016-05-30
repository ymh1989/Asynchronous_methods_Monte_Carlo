using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace MC_test
{
    public class Class1
    {
        [ExcelFunction(Description = "The generalized Black and Scholes formula (Closed-form)")]
        public static double GBS(int CallPutFlag, double S, double x, double T, double r, double b, double v)
        {
            double d1; double d2; double result = 0;
            d1 = (Math.Log(S / x) + (b + (v * v) / 2) * T) / (v * Math.Sqrt(T));
            d2 = d1 - v * Math.Sqrt(T);

            // Call
            if (CallPutFlag == 0) result = S * Math.Exp(-b * T) * CND(d1) - x * Math.Exp(-r * T) * CND(d2);
            // Put
            else result = result = x * Math.Exp(-r * T) * CND(-d2) - S * Math.Exp(-b * T) * CND(-d1);

            return result;
        }
        [ExcelFunction(Description = "The generalized Black and Scholes formula (Monte Carlo)")]
        public static double BSMC(int CallPutFlag, double S, double x, double T, double r, double b, double v, int Num_Simul)
        {
            int steps = 1;
            double[,] sp = new double[Num_Simul, steps + 1];
            double dt = T / steps;
            double payoff = 0;
            double result;

            double drift = ((r - b) - 0.5 * (v * v)) * dt;
            double sigsqrtdt = v * Math.Sqrt(dt);

            // Init
            for (int i = 0; i < Num_Simul; i++)
            {
                sp[i, 0] = S;
            }
            for (int i = 0; i < Num_Simul; i++)
            {
                for (int j = 1; j < steps + 1; j++)
                {
                    // System.Threading.Thread.Sleep(1);
                    sp[i, j] = sp[i, j - 1] * Math.Exp(drift + sigsqrtdt * randn());
                }
            }

            // Call
            if (CallPutFlag == 0)
            {
                for (int i = 0; i < Num_Simul; i++)
                {
                    payoff = payoff + Math.Max(sp[i, steps] - x, 0);
                }
            }
            // Put
            else
            {
                for (int i = 0; i < Num_Simul; i++)
                {
                    payoff = payoff + Math.Max(x - sp[i, steps], 0);
                }
            }
            result = Math.Exp(-r * T) * payoff / Num_Simul;


            return result;

        }


        /*Function**************************************************************************/
        public static double CND(double x)
        {
            double y; double Exponential; double SumA; double SumB;
            double result;
            y = Math.Abs(x);
            if (y > 37) result = 0;
            else
            {
                Exponential = Math.Exp(-(y * y) / 2);
                if (y < 7.07106781186547)
                {
                    SumA = 3.52624965998911E-02 * y + 0.700383064443688;
                    SumA = SumA * y + 6.37396220353165;
                    SumA = SumA * y + 33.912866078383;
                    SumA = SumA * y + 112.079291497871;
                    SumA = SumA * y + 221.213596169931;
                    SumA = SumA * y + 220.206867912376;
                    SumB = 8.83883476483184E-02 * y + 1.75566716318264;
                    SumB = SumB * y + 16.064177579207;
                    SumB = SumB * y + 86.7807322029461;
                    SumB = SumB * y + 296.564248779674;
                    SumB = SumB * y + 637.333633378831;
                    SumB = SumB * y + 793.826512519948;
                    SumB = SumB * y + 440.413735824752;
                    result = Exponential * SumA / SumB;
                }
                else
                {
                    SumA = y + 0.65;
                    SumA = y + 4 / SumA;
                    SumA = y + 3 / SumA;
                    SumA = y + 2 / SumA;
                    SumA = y + 1 / SumA;
                    result = Exponential / (SumA * 2.506628274631);
                }
            }
            if (x > 0)
            {
                result = 1 - result;
            }
            return result;
        }

        [ExcelFunction(Description = "Random Number Generator with Normal Distribution")]
        public static double randn()
        {
            double rand1, rand2;
            Random r = new Random(Guid.NewGuid().GetHashCode());

            rand1 = r.NextDouble();
            // return sqrt(variance * rand1) * sin(rand2);

            if (rand1 < 1e-100) rand1 = 1e-100;

            rand1 = -2 * Math.Log(1 - rand1);
            rand2 = r.NextDouble() * 2 * Math.PI;

            return Math.Sqrt(rand1) * Math.Cos(rand2);
        }
        /**************************************************************************/
    }
}

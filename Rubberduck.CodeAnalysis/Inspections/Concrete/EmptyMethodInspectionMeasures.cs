using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Inspections.Concrete.Extensions
{
    public static class EmptyMethodInspectionMeasures
    {
        public enum Syntax
        {
            Old,
            Linq
        }

        private static int count;
        private static double oldSum;
        private static double linqSum;

        public static void Measure(string what, Syntax syntax, int reps, Action action)
        {
            count += reps;
            action(); // warm up

            double[] results = new double[reps];
            double sum = 0;
            for (int i = 0; i < reps; ++i)
            {
                Stopwatch sw = Stopwatch.StartNew();
                action();
                results[i] = sw.Elapsed.TotalMilliseconds;
                sum += results[i];
            }
            Debug.Print("{0}\n{1} - AVG = {2}, Min = {3}, Max = {4}, Total = {5}",
                what, syntax.ToString(), results.Average(), results.Min(), results.Max(), sum);

            if (syntax == Syntax.Old) oldSum += sum;
            else linqSum += sum;
        }
    public static void DisplayResults()
    {
            Debug.Print("{0} - AVG = {1}, Total = {2}", Syntax.Old.ToString(), oldSum / count, oldSum);
            Debug.Print("{0} - AVG = {1}, Total = {2}", Syntax.Linq.ToString(), linqSum / count, linqSum);
        }

    }
}

using System;
using System.Collections.Generic;

namespace Rubberduck.UnitTesting
{
    public class TestResult
    {
        public TestResult(TestOutcome outcome, string output = "", long duration = 0)
        {
            Outcome = outcome;
            Output = output;
            Duration = duration;
        }

        public long Duration { get; private set; }
        public TestOutcome Outcome { get; private set; }
        public string Output { get; private set; }

        public override bool Equals(object obj)
        {
            if (obj is TestResult other)
            {
                return Outcome == other.Outcome
                    && Output == other.Output;
            }
            return false;
        }

        public override int GetHashCode()
        {
            var hashCode = -1268493841;
            hashCode = hashCode * -1521134295 + Outcome.GetHashCode();
            hashCode = hashCode * -1521134295 + EqualityComparer<string>.Default.GetHashCode(Output);
            return hashCode;
        }

        public override string ToString()
        {
            return $"{Outcome} ({Duration} ms) {Output}";
        }
    }
}

using Rubberduck.VBEditor;

namespace Rubberduck.UnitTesting
{
    public readonly struct TestResult
    {
        public TestResult(TestOutcome outcome, string output = "", long duration = 0)
        {
            Outcome = outcome;
            Output = output;
            Duration = duration;
        }

        public long Duration { get; }
        public TestOutcome Outcome { get; }
        public string Output { get; }

        public override int GetHashCode() => HashCode.Compute(Outcome, Output);
        public override string ToString() => $"{Outcome} ({Duration} ms) {Output}";
        public override bool Equals(object obj)
        {
            return obj is TestResult other
                    && Outcome == other.Outcome
                    && Output == other.Output;
        }
    }
}

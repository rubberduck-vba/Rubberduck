namespace Rubberduck.CodeAnalysis.CodeMetrics
{
    public enum AggregationLevel
    {
        Project = 1 << 0,
        Module = 1 << 1,
        Member = 1 << 2,
        Declaration = 1 << 3,
    }
}

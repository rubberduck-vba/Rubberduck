namespace Rubberduck.SmartIndenter
{
    internal enum AlignmentTokenType
    {
        Function,
        Parameter,
        Equals,
        Variable
    }

    internal class AlignmentToken
    {
        public AlignmentTokenType Type { get; }
        public int Position { get; }
        public int NestingDepth { get; }

        public AlignmentToken(AlignmentTokenType type, int position, int nesting = 0)
        {
            Type = type;
            Position = position;
            NestingDepth = nesting;
        }
    }
}

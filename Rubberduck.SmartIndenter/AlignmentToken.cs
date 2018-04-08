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

        public AlignmentToken(AlignmentTokenType type, int position)
        {
            Type = type;
            Position = position;
        }
    }
}

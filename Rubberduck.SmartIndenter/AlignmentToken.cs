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
        public AlignmentTokenType Type { get; private set; }
        public int Position { get; private set; }

        public AlignmentToken(AlignmentTokenType type, int position)
        {
            Type = type;
            Position = position;
        }
    }
}

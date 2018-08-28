namespace Rubberduck.AutoComplete.SelfClosingPairCompletion
{
    public class SelfClosingPair
    {
        public SelfClosingPair(char opening, char closing)
        {
            OpeningChar = opening;
            ClosingChar = closing;
        }

        public char OpeningChar { get; }
        public char ClosingChar { get; }
    }
}
namespace Rubberduck.AutoComplete.Service
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

        /// <summary>
        /// True if <see cref="OpeningChar"/> is the same as <see cref="ClosingChar"/>.
        /// </summary>
        public bool IsSymetric => OpeningChar == ClosingChar;
    }
}
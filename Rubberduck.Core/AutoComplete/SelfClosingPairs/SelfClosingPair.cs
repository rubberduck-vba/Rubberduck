using System;
using Rubberduck.VBEditor;

namespace Rubberduck.AutoComplete.SelfClosingPairs
{
    public class SelfClosingPair : IEquatable<SelfClosingPair>
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

        public bool Equals(SelfClosingPair other) => other?.OpeningChar == OpeningChar &&
                                                     other.ClosingChar == ClosingChar;

        public override bool Equals(object obj)
        {
            return obj is SelfClosingPair scp && Equals(scp);
        }

        public override int GetHashCode() => HashCode.Compute(OpeningChar, ClosingChar);
    }
}
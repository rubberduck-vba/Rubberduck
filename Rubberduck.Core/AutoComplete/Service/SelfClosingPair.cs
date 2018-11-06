using Rubberduck.VBEditor;
using System;

namespace Rubberduck.AutoComplete.Service
{
    public class SelfClosingPair : IEquatable<SelfClosingPair>
    {
        [Flags]
        public enum MatchType
        {
            NoMatch = 0,
            OpeningCharacterMatch = 1,
            ClosingCharacterMatch = 2,
        }

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
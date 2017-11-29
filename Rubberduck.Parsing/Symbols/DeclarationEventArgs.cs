using System;

namespace Rubberduck.Parsing.Symbols
{
    public class DeclarationEventArgs : EventArgs
    {
        public DeclarationEventArgs(Declaration declaration)
        {
            Declaration = declaration;
        }

        public Declaration Declaration { get; }
    }
}

using System;

namespace Rubberduck.Parsing.Symbols
{
    public class DeclarationEventArgs : EventArgs
    {
        private readonly Declaration _declaration;

        public DeclarationEventArgs(Declaration declaration)
        {
            _declaration = declaration;
        }

        public Declaration Declaration { get { return _declaration; } }
    }
}

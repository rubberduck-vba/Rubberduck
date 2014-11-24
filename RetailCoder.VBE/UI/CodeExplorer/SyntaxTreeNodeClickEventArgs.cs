using System;
using Rubberduck.VBA.Parser;

namespace Rubberduck.UI.CodeExplorer
{
    public class SyntaxTreeNodeClickEventArgs : EventArgs
    {
        public SyntaxTreeNodeClickEventArgs(Instruction instruction)
        {
            _instruction = instruction;
        }

        private readonly Instruction _instruction;
        public Instruction Instruction { get { return _instruction; } }
    }
}
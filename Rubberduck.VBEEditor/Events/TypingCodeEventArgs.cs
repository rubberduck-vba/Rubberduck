using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class TypingCodeEventArgs : EventArgs
    {
        public ICodePane CodePane { get; }

        public bool IsCommitted { get; }

        public string Code { get; }

        public TypingCodeEventArgs(ICodePane pane)
        {
            CodePane = pane;
            var selection = pane.Selection;
            using (var module = pane.CodeModule)
            {
                var atSelection = module.GetLines(selection);
                if (string.IsNullOrWhiteSpace(atSelection))
                {
                    IsCommitted = true;
                    Code = module.GetLines(selection.PreviousLine);
                }
                else
                {
                    Code = module.GetLines(selection);
                }
            }
        }
    }
}

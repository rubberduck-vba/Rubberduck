using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class AutoCompleteEventArgs : EventArgs
    {
        public AutoCompleteEventArgs(ICodePane pane)
        {
            CodePane = pane;
            var selection = pane.Selection;
            using (var module = pane.CodeModule)
            {
                ContentHash = module.ContentHash();
                var atSelection = module.GetLines(selection);
                if (string.IsNullOrWhiteSpace(atSelection))
                {
                    IsCommitted = true;
                    OldCode = module.GetLines(selection.PreviousLine);
                }
                else
                {
                    OldCode = module.GetLines(selection);
                }
            }
        }

        /// <summary>
        /// The CodePane wrapper for the module being edited.
        /// </summary>
        public ICodePane CodePane { get; }

        /// <summary>
        /// Indicates whether the line of code held in <see cref="OldCode"/> is committed or not.
        /// </summary>
        /// <remarks>
        /// If the line is committed, <see cref="OldCode"/> is located on the line that precedes the current selection in the <see cref="CodePane"/>.
        /// </remarks>
        public bool IsCommitted { get; }

        /// <summary>
        /// The content hash for the module before autocompletion. Used to prevent misfiring autocompletes.
        /// </summary>
        public string ContentHash { get; }

        /// <summary>
        /// If not committed, the entire current line of code. If committed, the line of code immediately preceding the current selection.
        /// </summary>
        public string OldCode { get; }

        /// <summary>
        /// The autocompleted line of code, assigned by the autocomplete implementation. Used for caching, to prevent misfiring autocompletes.
        /// If autocomplete works off committed input, this should match the <see cref="OldCode"/>.
        /// </summary>
        public string NewCode { get; set; }
    }
}

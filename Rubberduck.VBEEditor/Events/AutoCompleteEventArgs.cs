using System;
using System.Windows.Forms;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class AutoCompleteEventArgs : EventArgs
    {
        public AutoCompleteEventArgs(ICodeModule module, WindowsApi.KeyPressEventArgs e)
        {
            Character = e.Character;
            CodeModule = module;
            CurrentSelection = module.GetQualifiedSelection().Value.Selection;
            CurrentLine = module.GetLines(CurrentSelection);

            if (e.Key == Keys.Delete || 
                e.Key == Keys.Back ||
                e.Key == Keys.Enter)
            {
                Keys = e.Key;
            }
        }

        /// <summary>
        /// <c>true</c> if the character has been handled, i.e. written to the code pane.
        /// Set to <c>false</c> to swallow the character and prevent the WM message from reaching the code pane.
        /// </summary>
        public bool Handled { get; set; }

        /// <summary>
        /// The CodeModule wrapper for the module being edited.
        /// </summary>
        public ICodeModule CodeModule { get; }

        public bool IsCharacter => Keys == default;
        public char Character { get; }
        public Keys Keys { get; }

        public Selection CurrentSelection { get; }
        /// <summary>
        /// If not committed, the entire current line of code. If committed, the line of code immediately preceding the current selection.
        /// </summary>
        public string CurrentLine { get; }
    }
}

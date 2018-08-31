using System;
using System.Windows.Forms;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class AutoCompleteEventArgs : EventArgs
    {
        public AutoCompleteEventArgs(ICodeModule module, KeyPressEventArgs e)
        {
            if (e.Key == Keys.Delete ||
                e.Key == Keys.Back ||
                e.Key == Keys.Enter ||
                e.Key == Keys.Tab)
            {
                Keys = e.Key;
            }
            else
            {
                Character = e.Character;
            }
            CodeModule = module;
            CurrentSelection = module.GetQualifiedSelection().Value.Selection;
            CurrentLine = module.GetLines(CurrentSelection);
        }

        /// <summary>
        /// <c>true</c> if the character has been handled, i.e. written to the code pane.
        /// Set to <c>true</c> to swallow the character and prevent the WM message from reaching the code pane.
        /// </summary>
        public bool Handled { get; set; }

        /// <summary>
        /// The CodeModule wrapper for the module being edited.
        /// </summary>
        public ICodeModule CodeModule { get; }

        /// <summary>
        /// <c>true</c> if the event is originating from a <c>WM_CHAR</c> message.
        /// <c>false</c> if the event is originating from a <c>WM_KEYDOWN</c> message.
        /// </summary>
        /// <remarks>
        /// Inline completion is handled on WM_CHAR; deletions and block completion on WM_KEYDOWN.
        /// </remarks>
        public bool IsCharacter => Keys == default;
        /// <summary>
        /// The character whose key was pressed. Undefined value if <see cref="Keys"/> isn't `<see cref="Keys.None"/>.
        /// </summary>
        public char Character { get; }
        /// <summary>
        /// The actionnable key that was pressed. Value is <see cref="Keys.None"/> when <see cref="IsCharacter"/> is <c>true</c>.
        /// </summary>
        public Keys Keys { get; }

        /// <summary>
        /// The current location of the caret.
        /// </summary>
        public Selection CurrentSelection { get; }
        /// <summary>
        /// The contents of the current line of code.
        /// </summary>
        public string CurrentLine { get; }
    }
}

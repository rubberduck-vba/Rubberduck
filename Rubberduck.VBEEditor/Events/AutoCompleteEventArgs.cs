using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class AutoCompleteEventArgs : EventArgs
    {
        public AutoCompleteEventArgs(ICodeModule module, KeyPressEventArgs e)
        {
            Character = e.Character;
            CodeModule = module;
            CurrentSelection = module.GetQualifiedSelection().Value.Selection;
            CurrentLine = module.GetLines(CurrentSelection);
            ControlDown = e.ControlDown;
            IsDelete = e.IsDelete;
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
        /// The character whose key was pressed (Enter is always '\r'). Default value if Delete was pressed.
        /// </summary>
        public char Character { get; }

        /// <summary>
        /// <c>true</c> if the left control key was down on the keypress.
        /// </summary>
        public bool ControlDown { get; }

        /// <summary>
        /// <c>true</c> if the Delete key generated the event.
        /// </summary>
        public bool IsDelete { get; }
        
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

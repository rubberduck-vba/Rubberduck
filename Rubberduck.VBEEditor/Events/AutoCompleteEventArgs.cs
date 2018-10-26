using System;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.VBEditor.Events
{
    public class AutoCompleteEventArgs : EventArgs
    {
        public AutoCompleteEventArgs(ICodeModule module, KeyPressEventArgs e)
            : this(module, e.Character, e.ControlDown, e.IsDelete) { }

        public AutoCompleteEventArgs(ICodeModule module, char character, bool isControlKeyDown, bool isDeleteKey)
        {
            Module = module;
            IsControlKeyDown = isControlKeyDown;
            Character = character;
            IsDeleteKey = isDeleteKey;
        }

        public ICodeModule Module { get; }

        /// <summary>
        /// <c>true</c> if the character has been handled, i.e. written to the code pane.
        /// Set to <c>true</c> to swallow the character and prevent the WM message from reaching the code pane.
        /// </summary>
        public bool Handled { get; set; }

        /// <summary>
        /// The character whose key was pressed (Enter is always '\r'). Default value if Delete was pressed.
        /// </summary>
        public char Character { get; }

        /// <summary>
        /// <c>true</c> if the left control key was down on the keypress.
        /// </summary>
        public bool IsControlKeyDown { get; }

        /// <summary>
        /// <c>true</c> if the Delete key generated the event.
        /// </summary>
        public bool IsDeleteKey { get; }
    }
}

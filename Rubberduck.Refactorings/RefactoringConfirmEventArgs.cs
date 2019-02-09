using System;

namespace Rubberduck.Refactorings
{
    public class RefactoringConfirmEventArgs : EventArgs
    {
        public string Message { get; }
        public bool Confirm { get; set; }

        public RefactoringConfirmEventArgs(string message)
        {
            Message = message;
        }
    }
}
using System;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public class ActionFailedEventArgs : EventArgs
    {
        private readonly string _message;
        public string Message { get { return _message; } }

        private readonly string _title;
        public string Title { get { return _title; } }

        public ActionFailedEventArgs(string title, string message)
        {
            _title = title;
            _message = message;
        }

        public ActionFailedEventArgs(SourceControlException ex)
            : this(ex.Message, ex.InnerException.Message) { }
    }
}

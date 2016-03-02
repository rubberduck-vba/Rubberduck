using System;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public class ErrorEventArgs
    {
        public readonly string Message;
        public readonly string InnerMessage;

        public ErrorEventArgs(string message,string innerMessage)
        {
            Message = message;
            InnerMessage = innerMessage;
        }
    }

    public interface IControlViewModel
    {
        ISourceControlProvider Provider { get; set; }
        event EventHandler<ErrorEventArgs> ErrorThrown;
    }
}
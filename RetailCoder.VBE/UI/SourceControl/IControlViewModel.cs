using System;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public class ErrorEventArgs
    {
        public readonly string Message;

        public ErrorEventArgs(string message)
        {
            Message = message;
        }
    }

    public interface IControlViewModel
    {
        ISourceControlProvider Provider { get; set; }
        event EventHandler<ErrorEventArgs> ErrorThrown;
    }
}
using System;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    public class ErrorEventArgs
    {
        public readonly string Message;
        public readonly string InnerMessage;
        public readonly NotificationType NotificationType;

        public ErrorEventArgs(string message, string innerMessage, NotificationType notificationType)
        {
            Message = message;
            InnerMessage = innerMessage;
            NotificationType = notificationType;
        }
    }

    public interface IControlViewModel
    {
        ISourceControlProvider Provider { get; set; }
        event EventHandler<ErrorEventArgs> ErrorThrown;

        void RefreshView();
    }
}
using System;
using Rubberduck.SourceControl;

namespace Rubberduck.UI.SourceControl
{
    using System.Linq;

    public class ErrorEventArgs
    {
        public readonly string Title;
        public readonly string InnerMessage;
        public readonly NotificationType NotificationType;

        public ErrorEventArgs(string title, Exception innerException, NotificationType notificationType)
        {
            Title = title;
            InnerMessage = GetInnerExceptionMessage(innerException);
            NotificationType = notificationType;
        }

        public ErrorEventArgs(string title, string message, NotificationType notificationType)
        {
            Title = title;
            InnerMessage = message;
            NotificationType = notificationType;
        }

        private string GetInnerExceptionMessage(Exception ex)
        {
            return ex is AggregateException
                ? string.Join(Environment.NewLine, ((AggregateException) ex).InnerExceptions.Select(s => s.Message))
                : ex.Message;
        }
    }

    public interface IControlViewModel
    {
        SourceControlTab Tab { get; }

        ISourceControlProvider Provider { get; set; }
        event EventHandler<ErrorEventArgs> ErrorThrown;

        void RefreshView();
        void ResetView();
    }
}

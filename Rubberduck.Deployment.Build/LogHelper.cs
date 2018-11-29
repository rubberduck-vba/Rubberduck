using System;
using System.Diagnostics;
using Microsoft.Build.Framework;

namespace Rubberduck.Deployment.Build
{
    internal class RubberduckBuildEventArgs : CustomBuildEventArgs
    {
        internal RubberduckBuildEventArgs(string message, string helpKeyword, string senderName) : base(message, helpKeyword, senderName) { }
    }

    internal static class LogHelper
    {
        internal static void LogError(this ITask sender, Exception ex)
        {
            var stackTrace = new StackTrace(ex, true);
            var frame = stackTrace.GetFrame(0);
            var code = frame.GetMethod().Name;
            var file = frame.GetFileName();
            var lineNumber = frame.GetFileLineNumber();
            var columnNumber = frame.GetFileColumnNumber();
            var message = ex.Message;
            var helpKeyword = ex.HelpLink;
            var senderName = sender.GetType().FullName;

            var args = new BuildErrorEventArgs(
                subcategory: "Quack!",
                code: code,
                file: file,
                lineNumber: lineNumber,
                columnNumber: columnNumber,
                endLineNumber: lineNumber,
                endColumnNumber: columnNumber,
                message: message,
                helpKeyword: helpKeyword,
                senderName: senderName);

            sender.BuildEngine.LogErrorEvent(args);
        }

        internal static void LogCustom(this ITask sender, string message, string helpKeyword = null)
        {
            var senderName = sender.GetType().FullName;

            var args = new RubberduckBuildEventArgs(message, helpKeyword, senderName);

            sender.BuildEngine.LogCustomEvent(args);
        }
    }
}

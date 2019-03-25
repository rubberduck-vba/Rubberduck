using System;
using System.Diagnostics;
using Microsoft.Build.Framework;

namespace Rubberduck.Deployment.Build
{
    [Serializable]
    internal class RubberduckBuildEventArgs : BuildMessageEventArgs
    {
        // Default MSBuild verbosity is minimal; meaning that a normal message will not be shown.
        // Hence, we use high importance to ensure that it will still show in the build log at the
        // default verbosity level.
        internal RubberduckBuildEventArgs(string message, string helpKeyword, string senderName) : 
            base(message, helpKeyword, senderName, MessageImportance.High) { }
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

        internal static void LogMessage(this ITask sender, string message, string helpKeyword = null)
        {
            var senderName = sender.GetType().FullName;
            
            var args = new RubberduckBuildEventArgs(message, helpKeyword, senderName);
            sender.BuildEngine.LogMessageEvent(args);
        }
    }
}

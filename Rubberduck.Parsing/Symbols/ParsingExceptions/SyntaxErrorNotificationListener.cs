using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class SyntaxErrorNotificationListener : BaseErrorListener
    {
        public event EventHandler<SyntaxErrorEventArgs> OnSyntaxError;
        public override void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            var info = new SyntaxErrorInfo(msg, e, offendingSymbol, line, charPositionInLine);
            NotifySyntaxError(info);
        }

        private void NotifySyntaxError(SyntaxErrorInfo info)
        {
            OnSyntaxError?.Invoke(this, new SyntaxErrorEventArgs(info));
        }
    }

    public class SyntaxErrorEventArgs : EventArgs
    {
        public SyntaxErrorEventArgs(SyntaxErrorInfo info)
        {
            Info = info;
        }

        public SyntaxErrorInfo Info { get; }
    }
}
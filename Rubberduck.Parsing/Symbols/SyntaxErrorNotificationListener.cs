using System;
using Antlr4.Runtime;

namespace Rubberduck.Parsing.Symbols
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
            var handler = OnSyntaxError;
            if (handler != null)
            {
                handler.Invoke(this, new SyntaxErrorEventArgs(info));
            }
        }
    }

    public class SyntaxErrorEventArgs : EventArgs
    {
        private readonly SyntaxErrorInfo _info;

        public SyntaxErrorEventArgs(SyntaxErrorInfo info)
        {
            _info = info;
        }

        public SyntaxErrorInfo Info { get { return _info; } }
    }
}
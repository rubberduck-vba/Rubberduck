using System;
using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols.ParsingExceptions
{
    public class SyntaxErrorNotificationListener : BaseErrorListener
    {
        private readonly QualifiedModuleName _moduleName;

        public SyntaxErrorNotificationListener(QualifiedModuleName moduleName)
        {
            _moduleName = moduleName;
        }

        public event EventHandler<SyntaxErrorEventArgs> OnSyntaxError;
        public override void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string msg, RecognitionException e)
        {
            var info = new SyntaxErrorInfo(msg, e, offendingSymbol, line, charPositionInLine);
            NotifySyntaxError(info, _moduleName);
        }

        private void NotifySyntaxError(SyntaxErrorInfo info, QualifiedModuleName moduleName)
        {
            OnSyntaxError?.Invoke(this, new SyntaxErrorEventArgs(info, moduleName));
        }
    }

    public class SyntaxErrorEventArgs : EventArgs
    {
        public SyntaxErrorEventArgs(SyntaxErrorInfo info, QualifiedModuleName moduleName)
        {
            Info = info;
            ModuleName = moduleName;
        }

        public SyntaxErrorInfo Info { get; }

        public QualifiedModuleName ModuleName { get; }
    }
}
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.ParserErrors
{
    public class ParseErrorListItem
    {
        private readonly ParseErrorEventArgs _error;

        public ParseErrorListItem(ParseErrorEventArgs error)
        {
            _error = error;
        }

        public string ProjectName { get { return _error.ProjectName; } }
        public string ComponentName { get { return _error.ComponentName; } }
        public int ErrorLine { get { return _error.Exception.LineNumber; } }
        public int ErrorColumn { get { return _error.Exception.Position; } }
        public string ErrorToken { get { return _error.Exception.OffendingSymbol.Text; } }
        public string Message { get { return _error.Exception.Message; } }

        public string Value { get { return ToString(); } }

        public void Navigate()
        {
            _error.Navigate();
        }

        public override string ToString()
        {
            return string.Format("{0}.{1} ({2},{3}): {4}", ProjectName, ComponentName, ErrorLine, ErrorColumn, Message);
        }
    }
}

using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.UI.ParserErrors
{
    public class ParseErrorListItem
    {
        private readonly ParseErrorEventArgs _error;

        public ParseErrorListItem(ParseErrorEventArgs error)
        {
            _error = error;
        }

        public string ProjectName => _error.ProjectName;
        public string ComponentName => _error.ComponentName;
        public int ErrorLine => _error.Exception.LineNumber;
        public int ErrorColumn => _error.Exception.Position;
        public string ErrorToken => _error.Exception.OffendingSymbol.Text;
        public string Message => _error.Exception.Message;

        public string Value => ToString();

        public void Navigate(ISelectionService selectionService)
        {
            _error.Navigate(selectionService);
        }

        public override string ToString()
        {
            return string.Format("{0}.{1} ({2},{3}): {4}", ProjectName, ComponentName, ErrorLine, ErrorColumn, Message);
        }
    }
}

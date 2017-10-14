using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;

namespace Rubberduck.ToDoItems
{
    /// <summary>
    /// Represents a Todo comment and the necessary information to display and navigate to that comment.
    /// This is a binding item. Changing it's properties changes how it is displayed.
    /// </summary>
    public class ToDoItem : INavigateSource, IExportable
    {
        private readonly string _description;
        public string Description { get { return _description; } }

        private readonly string _type;
        public string Type { get { return _type; } }

        private readonly QualifiedSelection _selection;
        public QualifiedSelection Selection { get { return _selection; } }

        public ToDoItem(string markerText, CommentNode comment)
        {
            _description = comment.CommentText;
            _selection = comment.QualifiedSelection;
            _type = markerText;
        }

        public NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(_selection);
        }

        public object[] ToArray()
        {
            var module = _selection.QualifiedName;
            return new object[] { _type, Description, module.ProjectName, module.ComponentName, _selection.Selection.StartLine, _selection.Selection.StartColumn };
        }

        public string ToClipboardString()
        {
            var module = _selection.QualifiedName;
            return string.Format(
                RubberduckUI.ToDoExplorerToDoItemFormat,
                _type,
                _description,
                module.ProjectName,
                module.ComponentName,
                _selection.Selection.StartLine);
        }
    }
}

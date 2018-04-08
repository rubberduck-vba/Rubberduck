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
        public string Description { get; }

        public string Type { get; }

        public QualifiedSelection Selection { get; }

        public ToDoItem(string markerText, CommentNode comment)
        {
            Description = comment.CommentText;
            Selection = comment.QualifiedSelection;
            Type = markerText;
        }

        public NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(Selection);
        }

        public object[] ToArray()
        {
            var module = Selection.QualifiedName;
            return new object[] { Type, Description, module.ProjectName, module.ComponentName, Selection.Selection.StartLine, Selection.Selection.StartColumn };
        }

        public string ToClipboardString()
        {
            var module = Selection.QualifiedName;
            return string.Format(
                RubberduckUI.ToDoExplorerToDoItemFormat,
                Type,
                Description,
                module.ProjectName,
                module.ComponentName,
                Selection.Selection.StartLine);
        }
    }
}

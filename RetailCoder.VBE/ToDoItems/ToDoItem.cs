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
    public class ToDoItem : INavigateSource
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
    }
}

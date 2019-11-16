using Rubberduck.Interaction.Navigation;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.ToDoItems
{
    /// <summary>
    /// Represents a Todo comment and the necessary information to display and navigate to that comment.
    /// This is a binding item. Changing it's properties changes how it is displayed.
    /// </summary>
    public class ToDoItem : INavigateSource
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
    }
}

using Rubberduck.Parsing.Nodes;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.ToDoItems
{
    /// <summary>
    /// Represents a Todo comment and the necessary information to display and navigate to that comment.
    /// This is a binding item. Changing it's properties changes how it is displayed.
    /// </summary>
    public class ToDoItem
    {
        private readonly TodoPriority _priority;
        public TodoPriority Priority { get { return _priority; } }

        public string PriorityLabel { get { return RubberduckUI.ResourceManager.GetString("ToDoPriority_" + Priority, RubberduckUI.Culture); } }

        private readonly string _description;
        public string Description { get { return _description; } }

        private readonly string _projectName;
        public string ProjectName { get { return _projectName; } }

        private readonly string _moduleName;
        public string ModuleName { get { return _moduleName; } }

        private readonly int _lineNumber;
        public int LineNumber { get { return _lineNumber; } }

        private readonly QualifiedSelection _selection;
        public QualifiedSelection GetSelection() { return _selection; }

        public ToDoItem(TodoPriority priority, CommentNode comment)
            : this(priority, comment.CommentText, comment.QualifiedSelection)
        {
        }

        public ToDoItem(TodoPriority priority, string description, QualifiedSelection qualifiedSelection)
        {
            _priority = priority;
            _description = description;
            _selection = qualifiedSelection;
            _projectName = qualifiedSelection.QualifiedName.Project.Name;
            _moduleName = qualifiedSelection.QualifiedName.Component.Name;
            _lineNumber = qualifiedSelection.Selection.StartLine;
        }
    }
}

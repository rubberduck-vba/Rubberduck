using Rubberduck.Parsing;
using Rubberduck.Parsing.Nodes;

namespace Rubberduck.ToDoItems
{
    /// <summary>
    /// Represents a Todo comment and the necessary information to display and navigate to that comment.
    /// This is a binding item. Changing it's properties changes how it is displayed.
    /// </summary>
    public class ToDoItem
    {
        private readonly TaskPriority _priority;
        public TaskPriority Priority{ get { return _priority; } }

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

        public ToDoItem(TaskPriority priority, CommentNode comment)
            : this(priority, comment.CommentText, comment.QualifiedSelection)
        {
        }

        public ToDoItem(TaskPriority priority, string description, QualifiedSelection qualifiedSelection)
        {
            _priority = priority;
            _description = description;
            _selection = qualifiedSelection;
            _projectName = qualifiedSelection.QualifiedName.ProjectName;
            _moduleName = qualifiedSelection.QualifiedName.ModuleName;
            _lineNumber = qualifiedSelection.Selection.StartLine;
        }
    }
}

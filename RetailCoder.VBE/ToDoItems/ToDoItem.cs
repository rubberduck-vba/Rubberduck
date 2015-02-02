using System.Runtime.InteropServices;
using Rubberduck.Extensions;
using Rubberduck.VBA.Nodes;

namespace Rubberduck.ToDoItems
{
    [ComVisible(false)]
    public struct ToDoItem
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

        public ToDoItem(TaskPriority priority, CommentNode comment)
            : this(priority, comment.Comment, comment.QualifiedSelection)
        {
        }

        public ToDoItem(TaskPriority priority, string description, QualifiedSelection qualifiedSelection)
        {
            _priority = priority;
            _description = description;
            _projectName = qualifiedSelection.QualifiedName.ProjectName;
            _moduleName = qualifiedSelection.QualifiedName.ModuleName;
            _lineNumber = qualifiedSelection.Selection.StartLine;
        }
    }
}

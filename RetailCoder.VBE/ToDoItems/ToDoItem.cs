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

        private readonly QualifiedSelection _selection;
        public QualifiedSelection Selection { get { return _selection; } }

        public ToDoItem(TaskPriority priority, CommentNode comment)
            : this(priority, comment.Comment, comment.QualifiedSelection)
        {
        }

        public ToDoItem(TaskPriority priority, string description, QualifiedSelection qualifiedSelection)
        {
            _priority = priority;
            _description = description;
            _selection = qualifiedSelection;
        }
    }
}

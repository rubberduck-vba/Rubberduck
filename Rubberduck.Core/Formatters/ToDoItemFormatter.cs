using Rubberduck.ToDoItems;
using Rubberduck.Common;
using Rubberduck.Resources;

namespace Rubberduck.Formatters
{
    public class ToDoItemFormatter : IExportable
    {
        private readonly ToDoItem _toDoItem;

        public ToDoItemFormatter(ToDoItem toDoItem)
        {
            _toDoItem = toDoItem;
        }

        public object[] ToArray()
        {
            var module = _toDoItem.Selection.QualifiedName;
            return new object[] { _toDoItem.Type, _toDoItem.Description, module.ProjectName, module.ComponentName, _toDoItem.Selection.Selection.StartLine, _toDoItem.Selection.Selection.StartColumn };
        }

        public string ToClipboardString()
        {
            var module = _toDoItem.Selection.QualifiedName;
            return string.Format(RubberduckUI.ToDoExplorerToDoItemFormat,
                _toDoItem.Type,
                _toDoItem.Description,
                module.ProjectName,
                module.ComponentName,
                _toDoItem.Selection.Selection.StartLine);
        }
    }
}

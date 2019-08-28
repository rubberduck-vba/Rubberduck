using Rubberduck.ToDoItems;
using Rubberduck.Common;

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
            return _toDoItem.ToArray();
        }

        public string ToClipboardString()
        {
            return _toDoItem.ToClipboardString();
        }
    }
}

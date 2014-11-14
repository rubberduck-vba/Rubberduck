using Microsoft.Vbe.Interop;

namespace Rubberduck.ToDoItems
{
    internal enum TaskPriority
    {
        Low,
        Medium,
        High
    }

    internal class ToDoItem
    {
        public TaskPriority Priority{ get; set; }
        public string Description { get; set; }
        public string Module { get; set; } 
        public int LineNumber { get; set; }

        public ToDoItem(TaskPriority priority, string description, CodeModule module,  int lineNumber)
        {
            this.Priority = priority;
            this.Description = description.Trim();
            this.Module = module.Parent.Name;
            this.LineNumber = lineNumber;
        }
    }


}

using Microsoft.Vbe.Interop;

namespace Rubberduck.ToDoItems
{
    [System.Runtime.InteropServices.ComVisible(false)]
    public enum TaskPriority
    {
        Low,
        Medium,
        High
    }

    [System.Runtime.InteropServices.ComVisible(false)]
    public class ToDoItem
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

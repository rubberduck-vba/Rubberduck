using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RetailCoderVBE.Reflection;
using Microsoft.Vbe.Interop;

namespace RetailCoderVBE.TaskList
{
    internal enum TaskPriority
    {
        Low,
        Medium,
        High,
        Bug
    }

    internal class Task
    {
        public TaskPriority Priority{ get; set; }
        public string Description { get; set; }
        public CodeModule Module { get; set; } 
        public int LineNumber { get; set; }

        public Task(TaskPriority priority, string description, CodeModule module,  int lineNumber)
        {
            this.Priority = priority;
            this.Description = description;
            this.Module = module;
            this.LineNumber = lineNumber;
        }
    }


}

using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Rubberduck.VBA;
using Rubberduck.VBA.Grammar;

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

        // todo: 
        public ToDoItem(TaskPriority priority, ParserRuleContext context, string projectName, string moduleName)
            : this(priority, context.GetText(), projectName, moduleName, context.GetSelection().StartLine)
        {
        }

        public ToDoItem(TaskPriority priority, string description, string projectName, string moduleName, int lineNumber)
        {
            _priority = priority;
            _description = description;
            _projectName = projectName;
            _moduleName = moduleName;
            _lineNumber = lineNumber;
        }
    }
}

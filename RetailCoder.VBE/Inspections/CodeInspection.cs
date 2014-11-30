using System.Collections.Generic;
using System.Runtime.InteropServices;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public abstract class CodeInspection : IInspection
    {
        protected CodeInspection(string name, string message, CodeInspectionType type, CodeInspectionSeverity severity)
        {
            _name = name;
            _message = message;
            _inspectionType = type;
            Severity = severity;
        }

        private readonly string _name;
        public string Name { get { return _name; } }

        private readonly string _message;
        public string QuickFixMessage { get { return _message; } }

        private readonly CodeInspectionType _inspectionType;
        public CodeInspectionType InspectionType { get { return _inspectionType; } }

        public CodeInspectionSeverity Severity { get; set; }
        public bool IsEnabled { get; set; }

        /// <summary>
        /// Inspects specified tree node, searching for code issues.
        /// </summary>
        /// <param name="node"></param>
        /// <returns></returns>
        public abstract IEnumerable<CodeInspectionResultBase> Inspect(SyntaxTreeNode node);
    }
}
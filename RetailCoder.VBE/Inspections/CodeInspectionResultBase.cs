using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public abstract class CodeInspectionResultBase
    {
        public CodeInspectionResultBase(Instruction instruction, CodeInspectionSeverity type, string message)
        {
            _instruction = instruction;
            _type = type;
            _message = message;
        }

        private readonly Instruction _instruction;
        /// <summary>
        /// Gets the <see cref="Instruction"/> containing a code issue.
        /// </summary>
        public Instruction Instruction { get { return _instruction; } }

        private readonly CodeInspectionSeverity _type;
        /// <summary>
        /// Gets the severity of the code issue.
        /// </summary>
        public CodeInspectionSeverity Severity { get { return _type; } }

        private readonly string _message;
        /// <summary>
        /// Gets a short message that describes how the code issue can be fixed.
        /// </summary>
        public string Message { get { return _message; } }

        /// <summary>
        /// Addresses the issue by making changes to the code.
        /// </summary>
        /// <param name="vbe"></param>
        public abstract void QuickFix(VBE vbe);
    }
}

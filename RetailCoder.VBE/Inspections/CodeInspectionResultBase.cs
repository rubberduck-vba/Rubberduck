using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Parser;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public abstract class CodeInspectionResultBase
    {
        public CodeInspectionResultBase(string inspection, Instruction instruction, CodeInspectionSeverity type)
        {
            _name = inspection;
            _instruction = instruction;
            _type = type;
        }

        private readonly string _name;
        /// <summary>
        /// Gets a string containing the name of the code inspection.
        /// </summary>
        public string Name { get { return _name; } }

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

        /// <summary>
        /// Gets all available "quick fixes" for a code inspection result.
        /// </summary>
        /// <returns>Returns a <c>Dictionary&lt;string&gt;, Action&lt;VBE&gt;</c>
        /// where the keys are descriptions for each quick fix, and
        /// each value is a method returning <c>void</c> and taking a <c>VBE</c> parameter.</returns>
        public abstract IDictionary<string, Action<VBE>> GetQuickFixes();

        /// <summary>
        /// Gets/sets a value indicating whether inspection result has been handled/fixed.
        /// </summary>
        protected bool Handled { get; set; }

        /// <summary>
        /// Sets <see cref="Handled"/> flag to <c>true</c> without applying any fix.
        /// </summary>
        public void Ignore()
        {
            Handled = true;
        }
    }
}

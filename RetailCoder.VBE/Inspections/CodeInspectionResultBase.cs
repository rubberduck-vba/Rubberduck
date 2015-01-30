using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Antlr4.Runtime;
using Microsoft.Vbe.Interop;
using Rubberduck.VBA.Grammar;

namespace Rubberduck.Inspections
{
    [ComVisible(false)]
    public abstract class CodeInspectionResultBase
    {
        public CodeInspectionResultBase(string inspection, ParserRuleContext context, CodeInspectionSeverity type, string project, string module)
        {
            _name = inspection;
            _context = context;
            _type = type;
            _project = project;
            _module = module;
        }

        private readonly string _project;
        public string ProjectName { get { return _project; } }

        private readonly string _module;
        public string ModuleName { get { return _module; } }

        private readonly string _name;
        /// <summary>
        /// Gets a string containing the name of the code inspection.
        /// </summary>
        public string Name { get { return _name; } }

        private readonly ParserRuleContext _context;
        /// <summary>
        /// Gets the <see cref="Context"/> containing a code issue.
        /// </summary>
        public ParserRuleContext Context { get { return _context; } }

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
    }
}

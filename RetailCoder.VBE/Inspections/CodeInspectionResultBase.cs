using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.VBEInterfaces.RubberduckCodePane;

namespace Rubberduck.Inspections
{
    public abstract class CodeInspectionResultBase : ICodeInspectionResult
    {
        private readonly IRubberduckCodePaneFactory _factory;

        protected CodeInspectionResultBase(string inspection, CodeInspectionSeverity type, Declaration target, IRubberduckCodePaneFactory factory)
            : this(inspection, type, target.QualifiedName.QualifiedModuleName, null)
        {
            _target = target;
            _factory = factory;
        }

        /// <summary>
        /// Creates a comment inspection result.
        /// </summary>
        protected CodeInspectionResultBase(string inspection, CodeInspectionSeverity type, CommentNode comment)
            : this(inspection, type, comment.QualifiedSelection.QualifiedName, null, comment)
        { }

        /// <summary>
        /// Creates an inspection result.
        /// </summary>
        protected CodeInspectionResultBase(string inspection, CodeInspectionSeverity type, QualifiedModuleName qualifiedName, ParserRuleContext context, CommentNode comment = null)
        {
            _name = inspection;
            _type = type;
            _qualifiedName = qualifiedName;
            _context = context;
            _comment = comment;
        }

        private readonly string _name;
        /// <summary>
        /// Gets a string containing the name of the code inspection.
        /// </summary>
        public string Name { get { return _name; } }

        private readonly CodeInspectionSeverity _type;
        /// <summary>
        /// Gets the severity of the code issue.
        /// </summary>
        public CodeInspectionSeverity Severity { get { return _type; } }

        private readonly QualifiedModuleName _qualifiedName;
        protected QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private readonly ParserRuleContext _context;
        public ParserRuleContext Context { get { return _context; } }

        private readonly CommentNode _comment;
        public CommentNode Comment { get { return _comment; } }

        private readonly Declaration _target;
        protected Declaration Target { get { return _target; } }

        /// <summary>
        /// Gets the information needed to select the target instruction in the VBE.
        /// </summary>
        public virtual QualifiedSelection QualifiedSelection
        {
            get
            {
                if (_context == null && _comment == null)
                {
                    return _target.QualifiedSelection;
                }
                return _context == null
                    ? _comment.QualifiedSelection
                    : new QualifiedSelection(_qualifiedName, _context.GetSelection(), _factory);
            }
        }

        /// <summary>
        /// Gets all available "quick fixes" for a code inspection result.
        /// </summary>
        /// <returns>Returns a <c>Dictionary&lt;string&gt;, Action&lt;VBE&gt;</c>
        /// where the keys are descriptions for each quick fix, and
        /// each value is a parameterless method returning <c>void</c>.</returns>
        public abstract IDictionary<string, Action> GetQuickFixes();
    }
}

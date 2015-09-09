using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public abstract class CodeInspectionResultBase : ICodeInspectionResult
    {
        protected CodeInspectionResultBase(string inspection, CodeInspectionSeverity type, Declaration target)
            : this(inspection, type, target.QualifiedName.QualifiedModuleName, null)
        {
            _target = target;
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
        public QualifiedSelection QualifiedSelection
        {
            get
            {
                if (_context == null && _comment == null)
                {
                    return _target.QualifiedSelection;
                }
                return _context == null
                    ? _comment.QualifiedSelection
                    : new QualifiedSelection(_qualifiedName, _context.GetSelection());
            }
        }

        /// <summary>
        /// Gets all available "quick fixes" for a code inspection result.
        /// </summary>
        public virtual IEnumerable<CodeInspectionQuickFix> QuickFixes { get { return new CodeInspectionQuickFix[] {}; } }

        public bool HasQuickFixes { get { return QuickFixes.Any(); } }
    }
}

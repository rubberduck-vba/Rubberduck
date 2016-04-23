using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Nodes;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public abstract class InspectionResultBase : ICodeInspectionResult, INavigateSource
    {
        protected InspectionResultBase(IInspection inspection, Declaration target)
            : this(inspection, target.QualifiedName.QualifiedModuleName, target.Context)
        {
            _target = target;
        }

        /// <summary>
        /// Creates a comment inspection result.
        /// </summary>
        protected InspectionResultBase(IInspection inspection, CommentNode comment)
            : this(inspection, comment.QualifiedSelection.QualifiedName, null, comment)
        { }

        /// <summary>
        /// Creates an inspection result.
        /// </summary>
        protected InspectionResultBase(IInspection inspection, QualifiedModuleName qualifiedName, ParserRuleContext context, CommentNode comment = null)
        {
            _inspection = inspection;
            _qualifiedName = qualifiedName;
            _context = context;
            _comment = comment;
        }

        /// <summary>
        /// Creates an inspection result.
        /// </summary>
        protected InspectionResultBase(IInspection inspection, QualifiedModuleName qualifiedName, ParserRuleContext context, Declaration declaration, CommentNode comment = null)
        {
            _inspection = inspection;
            _qualifiedName = qualifiedName;
            _context = context;
            _target = declaration;
            _comment = comment;
        }

        private readonly IInspection _inspection;
        public IInspection Inspection { get { return _inspection; } }

        public abstract string Description { get; }

        private readonly QualifiedModuleName _qualifiedName;
        protected QualifiedModuleName QualifiedName { get { return _qualifiedName; } }

        private readonly ParserRuleContext _context;
        public ParserRuleContext Context { get { return _context; } }

        private readonly CommentNode _comment;
        public CommentNode Comment { get { return _comment; } }

        private readonly Declaration _target;
        protected virtual Declaration Target { get { return _target; } }

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

        public virtual CodeInspectionQuickFix DefaultQuickFix { get { return QuickFixes.FirstOrDefault(); } }

        public int CompareTo(ICodeInspectionResult other)
        {
            return Inspection.CompareTo(other.Inspection);
        }

        public override string ToString()
        {
            var module = QualifiedSelection.QualifiedName;
            return string.Format(
                InspectionsUI.QualifiedSelectionInspection,
                Inspection.Severity,
                Description,
                module.ProjectId,
                module.ComponentName,
                QualifiedSelection.Selection.StartLine);
        }

        public NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(QualifiedSelection);
        }

        public int CompareTo(object obj)
        {
            return CompareTo(obj as ICodeInspectionResult);
        }

        public object[] ToArray()
        {
            var module = QualifiedSelection.QualifiedName;
            return new object[] {Inspection.Severity.ToString(), Description, module.ProjectId, module.ComponentName, QualifiedSelection.Selection.StartLine };
        }

        public string ToCsvString()
        {
            var module = QualifiedSelection.QualifiedName;
            return string.Format(
                "\"{0}\",\"{1}\",\"{2}\",\"{3}\",{4}",
                Inspection.Severity,
                Description,
                module.ProjectId,
                module.ComponentName,
                QualifiedSelection.Selection.StartLine);
        }
    }
}

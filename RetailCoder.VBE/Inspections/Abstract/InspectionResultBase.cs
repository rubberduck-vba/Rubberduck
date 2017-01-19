using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class InspectionResultBase : IInspectionResult, INavigateSource
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
        public Declaration Target { get { return _target; } }

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
                    : new QualifiedSelection(_qualifiedName, _context.GetSelection());
            }
        }

        /// <summary>
        /// Gets all available "quick fixes" for a code inspection result.
        /// </summary>
        public virtual IEnumerable<QuickFixBase> QuickFixes { get { return Enumerable.Empty<QuickFixBase>(); } }

        public bool HasQuickFixes { get { return QuickFixes.Any(); } }

        public virtual QuickFixBase DefaultQuickFix { get { return QuickFixes.FirstOrDefault(); } }

        public virtual int CompareTo(IInspectionResult other)
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
                "(" + module.ProjectDisplayName + ")",
                module.ProjectName,
                module.ComponentName,
                QualifiedSelection.Selection.StartLine);
        }

        public virtual NavigateCodeEventArgs GetNavigationArgs()
        {
            return new NavigateCodeEventArgs(QualifiedSelection);
        }

        public int CompareTo(object obj)
        {
            return CompareTo(obj as IInspectionResult);
        }

        public object[] ToArray()
        {
            var module = QualifiedSelection.QualifiedName;
            return new object[] { Inspection.Severity.ToString(), module.ProjectName, module.ComponentName, Description, QualifiedSelection.Selection.StartLine, QualifiedSelection.Selection.StartColumn };
        }
    }
}

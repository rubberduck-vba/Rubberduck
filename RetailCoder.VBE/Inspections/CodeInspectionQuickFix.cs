using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public abstract class CodeInspectionQuickFix
    {
        private readonly ParserRuleContext _context;
        private readonly QualifiedSelection _selection;
        private readonly string _description;

        public CodeInspectionQuickFix(ParserRuleContext context, QualifiedSelection selection, string description)
        {
            _context = context;
            _selection = selection;
            _description = description;
        }

        public string Description { get { return _description; } }
        public ParserRuleContext Context { get { return _context; } }
        public QualifiedSelection Selection { get { return _selection; } }

        public abstract void Fix();

        /// <summary>
        /// Indicates whether this quickfix can be applied to all inspection results in module.
        /// </summary>
        /// <remarks>
        /// If both <see cref="CanFixInModule"/> and <see cref="CanFixInProject"/> are set to <c>false</c>,
        /// then the quickfix is only applicable to the current/selected inspection result.
        /// </remarks>
        public virtual bool CanFixInModule { get { return true; } }

        /// <summary>
        /// Indicates whether this quickfix can be applied to all inspection results in project.
        /// </summary>
        /// <remarks>
        /// If both <see cref="CanFixInModule"/> and <see cref="CanFixInProject"/> are set to <c>false</c>,
        /// then the quickfix is only applicable to the current/selected inspection result.
        /// </remarks>
        public virtual bool CanFixInProject { get { return true; } }
    }
}
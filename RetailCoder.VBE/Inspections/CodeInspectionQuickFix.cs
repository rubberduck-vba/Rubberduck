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
        protected ParserRuleContext Context { get { return _context; } }
        protected QualifiedSelection Selection { get { return _selection; } }

        public abstract void Fix();
    }
}
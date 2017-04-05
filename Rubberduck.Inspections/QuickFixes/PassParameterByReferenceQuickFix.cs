using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    /// <summary>
    /// Encapsulates a code inspection quickfix that changes a ByVal parameter into an explicit ByRef parameter.
    /// </summary>
    public class PassParameterByReferenceQuickFix : IQuickFix
    {
        private readonly IModuleRewriter _rewriter;
        private readonly IToken _token;

        public PassParameterByReferenceQuickFix(Declaration target, QualifiedSelection selection, IModuleRewriter rewriter)
            : base(target.Context, selection, InspectionsUI.PassParameterByReferenceQuickFix)
        {
            _rewriter = rewriter;
            _token = ((VBAParser.ArgContext)target.Context).BYVAL().Symbol;
        }

        public void Fix(IInspectionResult result)
        {
            _rewriter.Replace(_token, Tokens.ByRef);
        }
    }
}
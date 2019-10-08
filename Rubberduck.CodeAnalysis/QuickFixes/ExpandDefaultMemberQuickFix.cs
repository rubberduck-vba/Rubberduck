using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.QuickFixes
{
    public class ExpandDefaultMemberQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ExpandDefaultMemberQuickFix(IDeclarationFinderProvider declarationFinderProvider)
        : base(
            typeof(ObjectWhereProcedureIsRequiredInspection), 
            typeof(IndexedDefaultMemberAccessInspection), 
            typeof(IndexedRecursiveDefaultMemberAccessInspection), 
            typeof(ImplicitDefaultMemberAccessInspection), 
            typeof(ImplicitRecursiveDefaultMemberAccessInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);
            var finder = _declarationFinderProvider.DeclarationFinder;

            var lExpressionContext = result.Context;
            var selection = result.QualifiedSelection;
            InsertDefaultMember(lExpressionContext, selection, finder, rewriter);
        }

        private void InsertDefaultMember(ParserRuleContext lExpressionContext, QualifiedSelection selection, DeclarationFinder finder, IModuleRewriter rewriter)
        {
            var defaultMemberAccessCode = DefaultMemberAccessCode(selection, finder);
            rewriter.InsertAfter(lExpressionContext.Stop.TokenIndex, defaultMemberAccessCode);
        }

        private string DefaultMemberAccessCode(QualifiedSelection selection, DeclarationFinder finder)
        {
            var defaultMemberAccesses = finder.IdentifierReferences(selection).Where(reference => reference.DefaultMemberRecursionDepth > 0);
            var defaultMemberNames = defaultMemberAccesses.Select(reference => reference.Declaration.IdentifierName);
            return $".{string.Join("().", defaultMemberNames)}";
        }

        public override string Description(IInspectionResult result)
        {
            return Resources.Inspections.QuickFixes.ExpandDefaultMemberQuickFix;
        }

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
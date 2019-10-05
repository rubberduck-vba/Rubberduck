using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.QuickFixes
{
    public class ExpandBangNotationQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider; 

        public ExpandBangNotationQuickFix(IDeclarationFinderProvider declarationFinderProvider)
        : base(typeof(UseOfBangNotationInspection), typeof(UseOfRecursiveBangNotationInspection))
        {
            _declarationFinderProvider = declarationFinderProvider;
        }

        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(result.QualifiedSelection.QualifiedName);

            var dictionaryAccessContext = (VBAParser.DictionaryAccessContext)result.Context;
            AdjustArgument(dictionaryAccessContext, rewriter);

            var finder = _declarationFinderProvider.DeclarationFinder;
            var selection = result.QualifiedSelection;
            InsertDefaultMember(dictionaryAccessContext, selection, finder, rewriter);
        }

        private void AdjustArgument(VBAParser.DictionaryAccessContext dictionaryAccessContext, IModuleRewriter rewriter)
        {
            var argumentContext = ArgumentContext(dictionaryAccessContext);
            rewriter.InsertBefore(argumentContext.Start.TokenIndex, "(\"");
            rewriter.InsertAfter(argumentContext.Stop.TokenIndex, "\")");
        }

        private ParserRuleContext ArgumentContext(VBAParser.DictionaryAccessContext dictionaryAccessContext)
        {
            if (dictionaryAccessContext.parent is VBAParser.DictionaryAccessExprContext dictionaryAccess)
            {
                return dictionaryAccess.unrestrictedIdentifier();
            }

            return ((VBAParser.WithDictionaryAccessExprContext) dictionaryAccessContext.parent).unrestrictedIdentifier();
        }

        private void InsertDefaultMember(VBAParser.DictionaryAccessContext dictionaryAccessContext, QualifiedSelection selection, DeclarationFinder finder, IModuleRewriter rewriter)
        {
            var defaultMemberAccessCode = DefaultMemberAccessCode(selection, finder);
            rewriter.Replace(dictionaryAccessContext, defaultMemberAccessCode);
        }

        private string DefaultMemberAccessCode(QualifiedSelection selection, DeclarationFinder finder)
        {
            var defaultMemberAccesses = finder.IdentifierReferences(selection);
            var defaultMemberNames = defaultMemberAccesses.Select(reference => reference.Declaration.IdentifierName);
            return $".{string.Join("().", defaultMemberNames)}";
        }

        public override string Description(IInspectionResult result)
        {
            return Resources.Inspections.QuickFixes.ExpandBangNotationQuickFix;
        }

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;
    }
}
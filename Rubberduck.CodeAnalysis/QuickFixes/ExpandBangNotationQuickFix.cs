using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
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
            var defaultMemberNames = defaultMemberAccesses
                .Select(DefaultMemberName)
                .Select(declarationName => IsNotLegalIdentifierName(declarationName) 
                                            ? $"[{declarationName}]" 
                                            : declarationName);
            return $".{string.Join("().", defaultMemberNames)}";
        }

        private static string DefaultMemberName(IdentifierReference defaultMemberReference)
        {
            var defaultMemberMemberName = defaultMemberReference.Declaration.QualifiedName;
            var fullDefaultMemberName = $"{defaultMemberMemberName.QualifiedModuleName.ProjectName}.{defaultMemberMemberName.QualifiedModuleName.ComponentName}.{defaultMemberMemberName.MemberName}";

            if (DefaultMemberOverrides.TryGetValue(fullDefaultMemberName, out var defaultMemberOverride))
            {
                return defaultMemberOverride;
            }

            return defaultMemberMemberName.MemberName;
        }

        private bool IsNotLegalIdentifierName(string declarationName)
        {
            return string.IsNullOrEmpty(declarationName)
                || NonIdentifierCharacters.Any(character => declarationName.Contains(character))
                || AdditionalNonFirstIdentifierCharacters.Contains(declarationName[0]); ;
        }

        public override string Description(IInspectionResult result)
        {
            return Resources.Inspections.QuickFixes.ExpandBangNotationQuickFix;
        }

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;

        private string NonIdentifierCharacters = "[](){}\r\n\t.,'\"\\ |!@#$%^&*-+:=; ";
        private string AdditionalNonFirstIdentifierCharacters = "0123456789_";

        private static readonly Dictionary<string, string> DefaultMemberOverrides = new Dictionary<string, string>
        {
            ["Excel.Range._Default"] = "Item"
        };

    }
}
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
    public class ExpandDefaultMemberQuickFix : QuickFixBase
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        public ExpandDefaultMemberQuickFix(IDeclarationFinderProvider declarationFinderProvider)
        : base(
            typeof(ObjectWhereProcedureIsRequiredInspection), 
            typeof(IndexedDefaultMemberAccessInspection), 
            typeof(IndexedRecursiveDefaultMemberAccessInspection), 
            typeof(ImplicitDefaultMemberAccessInspection), 
            typeof(ImplicitRecursiveDefaultMemberAccessInspection),
            typeof(SuspiciousLetAssignmentInspection))
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

            if (result.Inspection is SuspiciousLetAssignmentInspection)
            {
                IdentifierReference rhsReference = result.Properties.RhSReference;
                var rhsLExpressionContext = rhsReference.Context;
                var rhsSelection = rhsReference.QualifiedSelection;
                InsertDefaultMember(rhsLExpressionContext, rhsSelection, finder, rewriter);
            }
        }

        private void InsertDefaultMember(ParserRuleContext lExpressionContext, QualifiedSelection selection, DeclarationFinder finder, IModuleRewriter rewriter)
        {
            var defaultMemberAccessCode = DefaultMemberAccessCode(selection, finder);
            rewriter.InsertAfter(lExpressionContext.Stop.TokenIndex, defaultMemberAccessCode);
        }

        private string DefaultMemberAccessCode(QualifiedSelection selection, DeclarationFinder finder)
        {
            var defaultMemberAccesses = finder.IdentifierReferences(selection).Where(reference => reference.DefaultMemberRecursionDepth > 0);
            var defaultMemberNames = defaultMemberAccesses.Select(DefaultMemberName)
                .Select(declarationName => IsNotLegalIdentifierName(declarationName)
                                            ? $"[{declarationName}]"
                                            : declarationName);
            return $".{string.Join("().", defaultMemberNames)}";
        }

        private bool IsNotLegalIdentifierName(string declarationName)
        {
            return string.IsNullOrEmpty(declarationName)
                || NonIdentifierCharacters.Any(character => declarationName.Contains(character))
                || AdditionalNonFirstIdentifierCharacters.Contains(declarationName[0]);                ;
        }

        private static string DefaultMemberName(IdentifierReference defaultMemberReference)
        {
            var defaultMemberMemberName = defaultMemberReference.Declaration.QualifiedName;
            var fullDefaultMemberName = $"{defaultMemberMemberName.QualifiedModuleName.ProjectName}.{defaultMemberMemberName.QualifiedModuleName.ComponentName}.{defaultMemberMemberName.MemberName}";

            if (DefaultMemberBaseOverrides.TryGetValue(fullDefaultMemberName, out var baseOverride))
            {
                if (DefaultMemberArgumentNumberOverrides.TryGetValue(fullDefaultMemberName, out var argumentNumberOverrides))
                {
                    var numberOfArguments = NumberOfArguments(defaultMemberReference);
                    if (argumentNumberOverrides.TryGetValue(numberOfArguments, out var numberOfArgumentsOverride))
                    {
                        return numberOfArgumentsOverride;
                    }
                }

                return baseOverride;
            }

            return defaultMemberMemberName.MemberName;
        }

        private static int NumberOfArguments(IdentifierReference defaultMemberReference)
        {
            if (defaultMemberReference.IsNonIndexedDefaultMemberAccess)
            {
                return 0;
            }

            var argumentList = ArgumentList(defaultMemberReference);
            if (argumentList == null)
            {
                return -1;
            }

            var arguments = argumentList.argument();

            return arguments?.Length ?? 0;
        }

        private static VBAParser.ArgumentListContext ArgumentList(IdentifierReference indexedDefaultMemberReference)
        {
            var defaultMemberReferenceContextWithArguments = indexedDefaultMemberReference.Context.Parent;
            switch (defaultMemberReferenceContextWithArguments)
            {
                case VBAParser.IndexExprContext indexExpression:
                    return indexExpression.argumentList();
                case VBAParser.WhitespaceIndexExprContext whiteSpaceIndexExpression:
                    return whiteSpaceIndexExpression.argumentList();
                default:
                    return null;
            }
        }

        public override string Description(IInspectionResult result)
        {
            return Resources.Inspections.QuickFixes.ExpandDefaultMemberQuickFix;
        }

        public override bool CanFixInProcedure => true;
        public override bool CanFixInModule => true;
        public override bool CanFixInProject => true;

        private string NonIdentifierCharacters = "[](){}\r\n\t.,'\"\\ |!@#$%^&*-+:=; ";
        private string AdditionalNonFirstIdentifierCharacters = "0123456789_";

        private static readonly Dictionary<string, string> DefaultMemberBaseOverrides = new Dictionary<string, string>
        {
            ["Excel.Range._Default"] = "Item"
        };

        private static readonly Dictionary<string, Dictionary<int, string>> DefaultMemberArgumentNumberOverrides = new Dictionary<string, Dictionary<int, string>>
        {
            ["Excel.Range._Default"] = new Dictionary<int, string>{[0] = "Value"}
        };
    }
}
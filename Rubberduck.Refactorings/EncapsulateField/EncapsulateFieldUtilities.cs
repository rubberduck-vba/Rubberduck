using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public struct EncapsulateFieldUtilities
    {
        public static bool IsModuleQualifiedExternalReferenceOfUDTField(IDeclarationFinderProvider declarationFinderProvider, IdentifierReference idRef, QualifiedModuleName qmnReferenceInstance)
        {
            if (idRef.QualifiedModuleName != qmnReferenceInstance)
            {
                if (qmnReferenceInstance.ComponentType == ComponentType.ClassModule)
                {
                    return true;
                }
                
                var module = declarationFinderProvider.DeclarationFinder.ModuleDeclaration(qmnReferenceInstance);
                return IsMemberAccessExpressionReference(module, idRef)
                    || IsWithMemberAccessExpressionReference(module, idRef);
            }

            return false;
        }

        public static bool IsRelatedUDTMemberReference(VariableDeclaration field, IdentifierReference udtMemberRef)
            => IsMemberAccessExpressionReference(field, udtMemberRef)
                || IsWithMemberAccessExpressionReference(field, udtMemberRef);

        private static bool IsMemberAccessExpressionReference(Declaration qualifyingDeclaration, IdentifierReference idRef)
        {
            if (qualifyingDeclaration is null || idRef is null)
            {
                return false;
            }

            var relevantRefs = qualifyingDeclaration.References.Where(rf => rf.Context.Parent is VBAParser.MemberAccessExprContext).ToList();
            if (!relevantRefs.Any())
            {
                return false;
            }

            foreach (var rf in relevantRefs)
            {
                var parent = rf.Context.Parent;
                while (parent is VBAParser.MemberAccessExprContext)
                {
                    if (parent == idRef.Context.Parent)
                    {
                        return true;
                    }
                    parent = parent.Parent;
                }
            }
            return false;
        }

        private static bool IsWithMemberAccessExpressionReference(Declaration qualifyingDeclaration, IdentifierReference idRef)
        {
            if (qualifyingDeclaration is null || idRef is null)
            {
                return false;
            }

            foreach (var rf in qualifyingDeclaration.References)
            {
                if (!rf.Context.TryGetAncestor<VBAParser.WithStmtContext>(out var qualifyingReferenceWithStmtContext))
                {
                    continue;
                }

                var withMemberAccessDescendentsOfQualifyingDeclaration = qualifyingReferenceWithStmtContext.GetDescendents<VBAParser.WithMemberAccessExprContext>().ToList();
                if (!withMemberAccessDescendentsOfQualifyingDeclaration.Any())
                {
                    continue;
                }

                if (withMemberAccessDescendentsOfQualifyingDeclaration.Contains(idRef.Context.Parent))
                {
                    return true;
                }

                if (idRef.Context.Parent is VBAParser.MemberAccessExprContext)
                {
                    var withMemberDescendentsOfReference = (idRef.Context.Parent as ParserRuleContext).GetDescendents<VBAParser.WithMemberAccessExprContext>();
                    var withMemberDescendentsInCommon = withMemberDescendentsOfReference.Intersect(withMemberAccessDescendentsOfQualifyingDeclaration);
                    if (withMemberDescendentsInCommon.Count() == 1)
                    {
                        return true;
                    }
                }
            }
            return false;
        }
    }
}

using System.Collections.Generic;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Abstract
{
    internal abstract class IdentifierReferenceInspectionFromDeclarationsBase : InspectionBase
    {
        protected IdentifierReferenceInspectionFromDeclarationsBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected abstract IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder);
        protected abstract string ResultDescription(IdentifierReference reference);

        protected virtual ICollection<string> DisabledQuickFixes(IdentifierReference reference) => new List<string>();

        /// <summary>
        /// Gets the possible <see cref="Declaration"/> that qualifies an identifier reference in a member access expression.
        /// </summary>
        protected IEnumerable<Declaration> GetQualifierCandidates(IdentifierReference reference, DeclarationFinder finder)
        {
            if (reference.Context.TryGetAncestor<VBAParser.CallStmtContext>(out var callStmt))
            {
                if (reference.Context.TryGetAncestor<VBAParser.LExpressionContext>(out var lExpression))
                {
                    // reference is in lexpression of a call statement

                    if (lExpression is VBAParser.MemberAccessExprContext member)
                    {
                        if (member.lExpression() is VBAParser.SimpleNameExprContext name)
                        {
                            if (reference.IdentifierName.Equals(name.identifier().GetText(), System.StringComparison.InvariantCultureIgnoreCase))
                            {
                                // unqualified
                                return Enumerable.Empty<Declaration>();
                            }

                            return finder.MatchName(name.identifier().GetText())
                                .Where(candidate => !candidate.Equals(reference.Declaration));
                        }

                        // todo get the actual qualifying declaration?
                        return finder.MatchName(member.lExpression().children.First().GetText());
                    }
                }

                
            }

            if (reference.Context.TryGetAncestor<VBAParser.MemberAccessExprContext>(out var memberAccess))
            {
                var parentModule = Declaration.GetModuleParent(reference.ParentScoping);
                var qualifyingExpression = memberAccess.lExpression();
                if (qualifyingExpression is VBAParser.SimpleNameExprContext simpleName)
                {
                    if (simpleName.GetText().Equals(Tokens.Me, System.StringComparison.InvariantCultureIgnoreCase))
                    {
                        // qualifier is 'Me'
                        return new[] { parentModule };
                    }

                    // todo get the actual qualifying declaration?
                    return finder.MatchName(simpleName.GetText());
                }
            }

            if (reference.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out var dot))
            {
                // qualifier is a With block
                var withBlock = dot.GetAncestor<VBAParser.WithStmtContext>();
                return finder.ContainedIdentifierReferences(new QualifiedSelection(reference.QualifiedModuleName, withBlock.GetSelection()))
                    .Select(r => r.Declaration).Distinct();
            }

            return Enumerable.Empty<Declaration>();
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            var objectionableReferences = ObjectionableReferences(finder);
            var resultReferences = ResultReferences(objectionableReferences, finder);
            return resultReferences
                .Select(reference => InspectionResult(reference, finder))
                .ToList();
        }

        private IEnumerable<IdentifierReference> ResultReferences(IEnumerable<IdentifierReference> potentialResultReferences, DeclarationFinder finder)
        {
            return potentialResultReferences
                .Where(reference => IsResultReference(reference, finder));
        }

        protected virtual IEnumerable<IdentifierReference> ObjectionableReferences(DeclarationFinder finder)
        {
            var objectionableDeclarations = ObjectionableDeclarations(finder);
            return objectionableDeclarations
                .SelectMany(declaration => declaration.References);
        }

        protected virtual bool IsResultReference(IdentifierReference reference, DeclarationFinder finder) => true;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableReferences = ObjectionableReferences(finder)
                .Where(reference => reference.QualifiedModuleName.Equals(module));
            var resultReferences = ResultReferences(objectionableReferences, finder);
            return resultReferences
                .Select(reference => InspectionResult(reference, finder))
                .ToList();
        }

        protected virtual IInspectionResult InspectionResult(IdentifierReference reference, DeclarationFinder finder)
        {
            return new IdentifierReferenceInspectionResult(
                this,
                ResultDescription(reference),
                finder,
                reference,
                DisabledQuickFixes(reference));
        }
    }

    internal abstract class IdentifierReferenceInspectionFromDeclarationsBase<T> : InspectionBase
    {
        protected IdentifierReferenceInspectionFromDeclarationsBase(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider)
        {}

        protected abstract IEnumerable<Declaration> ObjectionableDeclarations(DeclarationFinder finder);
        protected abstract (bool isResult, T properties) IsResultReferenceWithAdditionalProperties(IdentifierReference reference, DeclarationFinder finder);
        protected abstract string ResultDescription(IdentifierReference reference, T properties);

        protected virtual ICollection<string> DisabledQuickFixes(IdentifierReference reference, T properties) => new List<string>();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            var objectionableReferences = ObjectionableReferences(finder);
            var resultReferences = ResultReferences(objectionableReferences, finder);
            return resultReferences
                .Select(tpl => InspectionResult(tpl.reference, finder, tpl.properties))
                .ToList();
        }

        private IEnumerable<(IdentifierReference reference, T properties)> ResultReferences(IEnumerable<IdentifierReference> potentialResultReferences, DeclarationFinder finder)
        {
            return potentialResultReferences
                .Select(reference => ReferenceWithResultProperties(reference, finder))
                .Where(result => result.HasValue)
                .Select(result => result.Value); ;
        }

        private (IdentifierReference reference, T properties)? ReferenceWithResultProperties(IdentifierReference reference, DeclarationFinder finder)
        {
            var (isResult, properties) = IsResultReferenceWithAdditionalProperties(reference, finder);
            return isResult
                ? (reference, properties)
                : ((IdentifierReference reference, T properties)?)null;
        }

        protected virtual IEnumerable<IdentifierReference> ObjectionableReferences(DeclarationFinder finder)
        {
            var objectionableDeclarations = ObjectionableDeclarations(finder);
            return objectionableDeclarations
                .SelectMany(declaration => declaration.References);
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var objectionableReferences = ObjectionableReferences(finder)
                .Where(reference => reference.QualifiedModuleName.Equals(module));
            var resultReferences = ResultReferences(objectionableReferences, finder);
            return resultReferences
                .Select(tpl => InspectionResult(tpl.reference, finder, tpl.properties))
                .ToList();
        }

        protected virtual IInspectionResult InspectionResult(IdentifierReference reference, DeclarationFinder finder, T properties)
        {
            return new IdentifierReferenceInspectionResult<T>(
                this,
                ResultDescription(reference, properties),
                finder,
                reference,
                properties,
                DisabledQuickFixes(reference, properties));
        }
    }
}
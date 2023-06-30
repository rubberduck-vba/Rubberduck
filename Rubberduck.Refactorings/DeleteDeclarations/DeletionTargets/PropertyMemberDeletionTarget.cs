using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Refactorings.DeleteDeclarations
{
    internal class PropertyMemberDeletionTarget : ModuleElementDeletionTarget, IPropertyDeletionTarget
    {
        public PropertyMemberDeletionTarget(IDeclarationFinderProvider declarationFinderProvider, Declaration target, IModuleRewriter rewriter)
            : base(declarationFinderProvider, target, rewriter)
        { }

        public override string BuildEOSReplacementContent()
        {
            if (ModifiedTargetEOSContent.Contains(Tokens.CommentMarker))
            {
                return base.BuildEOSReplacementContent();
            }

            var body = GetCurrentTextPriorToSeparationAndIndentation(PrecedingEOSContext, Rewriter);

            var ending = IsGroupedWithRelatedProperties() && IsLastPropertyOfGroup()
                ? $"{EOSSeparation}{EOSIndentation}"
                : $"{PrecedingEOSSeparation}{EOSIndentation}";

            return body + ending;
        }

        public bool IsGroupedWithRelatedProperties()
        {
            var propertiesOfSameName = DeclarationFinderProvider.DeclarationFinder
                .MatchName(TargetProxy.IdentifierName)
                .ToList();

            if (propertiesOfSameName?.Count == 1)
            {
                return false;
            }

            var orderedPropertiesOfSameName = propertiesOfSameName
                .OrderBy(d => d.Selection)
                .ToList();

            var allOtherMembers = DeclarationFinderProvider.DeclarationFinder
                .Members(TargetProxy.QualifiedModuleName)
                .Where(d => d.DeclarationType.HasFlag(DeclarationType.Member) && !propertiesOfSameName.Contains(d));

            bool DeclarationsAreAdjacent(Declaration first, Declaration next)
                => !allOtherMembers.Any(proc => proc.Selection > first.Selection && proc.Selection < next.Selection);

            var grouped = new List<Declaration>();

            for (var position = 0; position + 1 < orderedPropertiesOfSameName.Count; position++)
            {
                var first = orderedPropertiesOfSameName.ElementAt(position);
                var second = orderedPropertiesOfSameName.ElementAt(position + 1);

                if ((first == TargetProxy || second == TargetProxy) && DeclarationsAreAdjacent(first, second))
                {
                    grouped.Add(first);
                    grouped.Add(second);
                }
            }

            grouped.Distinct();
            return grouped.Count > 1;
        }

        private bool IsLastPropertyOfGroup()
        {
            var lastPropertyAcessorDeclaration = DeclarationFinderProvider.DeclarationFinder
                .MatchName(TargetProxy.IdentifierName)
                .OrderBy(d => d.Selection).LastOrDefault();

            return lastPropertyAcessorDeclaration == TargetProxy;
        }
    }
}

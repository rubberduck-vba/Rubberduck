using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IUDTMemberReferenceProvider
    {
        /// <summary>
        /// Returns an IDictionary interface mapping a UserDefinedType field to references of it's UDTMembers. 
        /// </summary>
        IDictionary<Declaration, List<IdentifierReference>> UdtFieldToMemberReferences(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<VariableDeclaration> selected, Func<IEnumerable<Declaration>, bool> udtMemberIsUDTPredicate = null);
    }

    public class UDTMemberReferenceProvider : IUDTMemberReferenceProvider
    {
        public IDictionary<Declaration, List<IdentifierReference>> UdtFieldToMemberReferences(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<VariableDeclaration> targetFields, Func<IEnumerable<Declaration>, bool> udtMemberIsUDTPredicate = null)
        {
            var udtFieldToMemberReferences = new Dictionary<Declaration, List<IdentifierReference>>();
            var relevantReferences = new List<IdentifierReference>();
            foreach (var udtField in targetFields.Where(s => s.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false))
            {
                var lowestLeafUdtMembers = FlattenNestedUdtMembers(declarationFinderProvider, udtField, udtMemberIsUDTPredicate ?? UdtMemberIsUDTPredicateDefault);
                    
                var udtRefs = lowestLeafUdtMembers.SelectMany(d => d.References)
                    .Where(rf => EncapsulateFieldUtilities.IsRelatedUDTMemberReference(udtField, rf))
                    .Select(rf => rf);

                relevantReferences.AddRange(udtRefs);
                udtFieldToMemberReferences.Add(udtField, udtRefs.ToList());
            }

            return udtFieldToMemberReferences;
        }

        private static List<Declaration> FlattenNestedUdtMembers(IDeclarationFinderProvider declarationFinderProvider, VariableDeclaration target, Func<List<Declaration>, bool> containsUDT)
        {
            var udtMembers = declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                .Where(udtm => udtm.ParentDeclaration == target.AsTypeDeclaration).ToList();

            const int maxNestedUDTIterations = 20;
            var guard = 0;
            while(containsUDT(udtMembers))
            {
                if (guard++ >= maxNestedUDTIterations)
                {
                    throw new ArgumentException("Unable to resolve UDT Members for 'target' argument");
                }

                var UDTs = udtMembers.Where(d => d.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false);

                var childrenUDTMembers = UDTs
                    .SelectMany(u => declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.UserDefinedTypeMember)
                        .Where(udtm => udtm.ParentDeclaration == u.AsTypeDeclaration))
                    .ToList();

                udtMembers.RemoveAll(d => UDTs.Contains(d));
                udtMembers.AddRange(childrenUDTMembers);
            }
            return udtMembers;
        }
        private static bool UdtMemberIsUDTPredicateDefault(IEnumerable<Declaration> udtMembers)
            => udtMembers.Any(d => d.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.UserDefinedType) ?? false);
    }
}
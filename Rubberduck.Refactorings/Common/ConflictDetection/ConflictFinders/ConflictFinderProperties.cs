using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Finds identifier conflicts for Let\Set\Get Properties
    /// </summary>
    /// <seealso cref="ConflictFinderMembers"/>
    public class ConflictFinderProperties : ConflictFinderModuleCodeSection
    {
        public ConflictFinderProperties(IDeclarationFinderProvider declarationFinderProvider, IDeclarationProxyFactory proxyFactory)
            : base(declarationFinderProvider, proxyFactory) { }

        public override bool TryFindConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            if (proxy.DeclarationType.Equals(DeclarationType.Property))
            {
                return TryFindPropertyTypesConflicts(proxy, sessionData, out conflicts);
            }
            return TryFindPropertyNameConflict(proxy, sessionData, out conflicts);
        }

        private bool TryFindPropertyTypesConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            var result = false;
            try
            {
                result =  TryFindPropertyTypesConflictsProtected(proxy, sessionData, out conflicts);
            }
            catch { }

            proxy.DeclarationType = DeclarationType.Property;
            return result;
        }

        private bool TryFindPropertyTypesConflictsProtected(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();
            var propertyTypes = new DeclarationType[]
            {
                DeclarationType.PropertyGet,
                DeclarationType.PropertyLet,
                DeclarationType.PropertySet
            };

            foreach (var propertyType in propertyTypes)
            {
                proxy.DeclarationType = propertyType;
                if (TryFindPropertyNameConflict(proxy, sessionData, out var propertyConflicts))
                {
                    conflicts = AddConflicts(conflicts, proxy, propertyConflicts);
                }
            }
            return conflicts.Values.Any();
        }

        //MS-VBAL 5.3.1.7
        //Each property declaration must have a procedure name that is different from the 
        //name of any other module variable, module constant, enum member name, 
        //external procedure, function, or subroutine that is defined within the same module.
        private bool TryFindPropertyNameConflict(IConflictDetectionDeclarationProxy memberProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            var allMatches = IdentifierMatches(memberProxy, sessionData, out var targetModuleMatches);

            var targetModulePropertyIdentifierMatches = targetModuleMatches.Where(idm => idm.DeclarationType.HasFlag(DeclarationType.Property));
            if (targetModulePropertyIdentifierMatches.Any())
            {
                //5.3.1.7 
                //Each property Get must have a unique name
                //Each property Let must have a unique name
                //Each property Set must have a unique name
                var propertyIdentifierMatches = targetModulePropertyIdentifierMatches.Where(p => p.DeclarationType == memberProxy.DeclarationType
                                                                    || p.DeclarationType.Equals(DeclarationType.Property)
                                                                    || p.Prototype == null && memberProxy.Prototype == null);
                if (propertyIdentifierMatches.Any())
                {
                    conflicts = AddConflicts(conflicts, memberProxy, propertyIdentifierMatches);
                    return true;
                }

                //5.3.1.7 each property that shares a common name must have equivalent parameter lists
                if (memberProxy.Prototype != null
                    && !HaveEquivalentParameterLists(memberProxy.Prototype, targetModulePropertyIdentifierMatches.First().Prototype))
                {
                    conflicts = AddConflicts(conflicts, memberProxy, targetModulePropertyIdentifierMatches);
                }
            }

            if (TryFindMemberConflictChecksCommon(memberProxy, targetModuleMatches.Except(targetModulePropertyIdentifierMatches), out var commonChecks))
            {
                conflicts = AddConflicts(conflicts, memberProxy, commonChecks);
            }

            if (TryFindDeclarationConflictWithOtherNonModuleQualifiedReferences(memberProxy, allMatches, out var conflictRefs3))
            {
                conflicts = AddReferenceConflicts(conflicts, sessionData, memberProxy, conflictRefs3);
            }

            if (memberProxy.HasStandardModuleParent)
            {
                if (NonModuleQualifiedReferenceConflicts(memberProxy, sessionData, allMatches, out var nonModuleQualifiedRefDeclarations))
                {
                    conflicts = AddConflicts(conflicts, memberProxy, nonModuleQualifiedRefDeclarations);
                }
            }

            return conflicts.Values.Any();
        }

        private static bool HaveEquivalentParameterLists(Declaration proxyDeclaration, Declaration existingProperty)
        {
            var propertyAsType = GetPropertyAsTypeName(existingProperty);
            var proxyAsType = GetPropertyAsTypeName(proxyDeclaration);

            if (!propertyAsType.Equals(proxyAsType))
            {
                return false;
            }

            var propertyParamsToEvaluate = GetPropertyParameters(existingProperty);

            var proxyParamsToEvaluate = GetPropertyParameters(proxyDeclaration);

            if (propertyParamsToEvaluate.Count() != proxyParamsToEvaluate.Count())
            {
                return false;
            }

            for (var idx = 0; idx < propertyParamsToEvaluate.Count(); idx++)
            {
                var propertyParam = propertyParamsToEvaluate.ElementAt(idx);
                var proxyParam = proxyParamsToEvaluate.ElementAt(idx);

                if (proxyParam.AsTypeName != propertyParam.AsTypeName)
                {
                    return false;
                }

                if (!UsesEquivalentParameterMechanism(propertyParam, proxyParam))
                {
                    return false;
                }

                if (propertyParam.IdentifierName != proxyParam.IdentifierName)
                {
                    return false;
                }
                //Note: MS-VBAL indicates that the number of Optional parameters must match.  
                //However, no scenario was found that could get the VBE to identify non-compilable code.
                //So, no checks are added for that condition.

                //This applies only be the last parameter (except the RHS value of a Get) - but there is no harm in checking them all
                if (propertyParam.IsParamArray != proxyParam.IsParamArray)
                {
                    return false;
                }
            }
            return true;
        }

        private static string GetPropertyAsTypeName(Declaration declaration)
        {
            Debug.Assert(declaration.DeclarationType.HasFlag(DeclarationType.Property));

            if (declaration is IParameterizedDeclaration pDec
                && !declaration.DeclarationType.Equals(DeclarationType.PropertyGet))
            {
                return pDec.Parameters.Last().AsTypeName;
            }
            return declaration.AsTypeName;
        }

        private static IReadOnlyList<ParameterDeclaration> GetPropertyParameters(Declaration declaration)
        {
            Debug.Assert(declaration.DeclarationType.HasFlag(DeclarationType.Property));

            if (declaration is IParameterizedDeclaration pDec)
            {
                return !declaration.DeclarationType.Equals(DeclarationType.PropertyGet)
                    ? pDec.Parameters.Take(pDec.Parameters.Count - 1).ToList()
                    : pDec.Parameters;
            }
            return new List<ParameterDeclaration>();
        }

        private static bool UsesEquivalentParameterMechanism(ParameterDeclaration existingParam, ParameterDeclaration proxyParam)
        {
            var proxyIsByRef = (proxyParam.IsImplicitByRef || proxyParam.IsImplicitByRef);
            if (existingParam.IsImplicitByRef || existingParam.IsByRef)
            {
                return proxyIsByRef;
            }
            //The existing parameter is ByVal
            return !proxyIsByRef;
        }
    }
}

using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    public class ConflictFinderProperties : ConflictFinderModuleCodeSection
    {
        public ConflictFinderProperties(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        public override bool TryFindConflict(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            if (proxy.DeclarationType.Equals(DeclarationType.Property))
            {
                return TryFindPropertyTypeConflicts(proxy, sessionData, out conflicts);
            }
            return TryFindPropertyNameConflict(proxy, sessionData, out conflicts);
        }

        private bool TryFindPropertyTypeConflicts(IConflictDetectionDeclarationProxy proxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
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

            if (!IsExistingTargetModule(memberProxy, out var targetModule))
            {
                return false;
            }

            var allMatches = IdentifierMatches(memberProxy, sessionData, out var targetModuleMatches);

            var targetModulePropertyIdentifierMatches = targetModuleMatches.Where(idm => idm.DeclarationType.HasFlag(DeclarationType.Property));
            if (targetModulePropertyIdentifierMatches.Any())
            {
                //5.3.1.7 
                //Each property Get must have a unique name
                //Each property Let must have a unique name
                //Each property Set must have a unique name
                if (targetModulePropertyIdentifierMatches.Any(p => p.DeclarationType == memberProxy.DeclarationType))
                {
                    conflicts = AddConflicts(conflicts, memberProxy, targetModulePropertyIdentifierMatches.Where(p => p.DeclarationType == memberProxy.DeclarationType));
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

            if (NonModuleQualifiedMemberReferenceConflicts(memberProxy, sessionData, allMatches, out var nonModuleQualifiedRefDeclarations))
            {
                conflicts = AddConflicts(conflicts, memberProxy, nonModuleQualifiedRefDeclarations);
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
                //However, no scenario was found that could get the VBE to complain.
                //So, no checks are added for that condition.

                //This can only be the last parameter (except the RHS value of a Get) - but there is no harm in checking them all
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

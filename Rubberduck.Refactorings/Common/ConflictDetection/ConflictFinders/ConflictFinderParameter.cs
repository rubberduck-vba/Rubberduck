using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Common
{
    /// <summary>
    /// Finds identifier conflicts for Parameters
    /// </summary>
    /// <remarks>
    ///MS-VBAL 5.3.1.5
    ///1. The name of each<positional-param>, <optional-param>, and<param-array> that are elements of a 
    ///function declaration must be different from the name of the function declaration.
    ///2. Each<positional-param>, <optional-param>, and<param-array> that are elements of the 
    ///same<parameter-list>, <property-parameters>, or<event-parameter-list> must have a distinct names. 
    ///3. The name value of a<positional-param>, <optional-param>, or a<param-array> may not be the same
    ///as the name of any variable defined by a <dim-statement>, a<redim-statement>,
    ///or a<const-statement> within the<procedure-body> of the containing procedure declaration.

    /// </remarks>
    class ConflictFinderParameter : ConflictFinderBase
    {
        public ConflictFinderParameter(IDeclarationFinderProvider declarationFinderProvider)
            : base(declarationFinderProvider) { }

        public override bool TryFindConflict(IConflictDetectionDeclarationProxy parameterProxy, IConflictDetectionSessionData sessionData, out Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>> conflicts)
        {
            conflicts = new Dictionary<IConflictDetectionDeclarationProxy, List<IConflictDetectionDeclarationProxy>>();

            /*1 Different than function*/
            if (parameterProxy.Prototype != null && parameterProxy.Prototype.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Function)
                && AreVBAEquivalent(parameterProxy.IdentifierName, parameterProxy.Prototype.ParentDeclaration.IdentifierName))
            {
                conflicts = AddConflicts(conflicts, parameterProxy, sessionData.CreateProxy(parameterProxy.Prototype.ParentDeclaration));
            }

            /*2 Unique params*/
            var memberScopeMatches = _declarationFinderProvider.DeclarationFinder.MatchName(parameterProxy.IdentifierName)
                                                                        .Where(d => d.ParentScopeDeclaration.IsMember()
                                                                                && d.QualifiedModuleName.ComponentName == parameterProxy.TargetModuleName
                                                                                && d.ProjectId == (parameterProxy.ProjectId))
                                                                        .Select(m => sessionData.CreateProxy(m));
            if (memberScopeMatches.Any())
            {
                conflicts = AddConflicts(conflicts, parameterProxy, memberScopeMatches);
            }

            /*3 Parameter is different than parent procedure references within the body (i.e., recursive calls) .
            *Strictly speaking, this exceeds 5.3.1.5 (#3 above).  However, changing a parameter
            *to match any referenced element within the procedure body will generate either uncompilable
            *code or change the resulting logic of the procedure.  Flag it as a conflict. */
            var procedureBodyReferences = _declarationFinderProvider.DeclarationFinder.IdentifierReferences(parameterProxy.Prototype.ParentDeclaration.QualifiedName)
                    .Where(rf => AreVBAEquivalent(rf.IdentifierName, parameterProxy.IdentifierName));

            if (procedureBodyReferences.Any())
            {
                conflicts = AddReferenceConflicts(conflicts, sessionData, parameterProxy, procedureBodyReferences);
            }

            return conflicts.Values.Any();
        }

    }
}

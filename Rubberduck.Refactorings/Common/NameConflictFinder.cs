using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.Common
{
    public interface INameConflictFinder
    {
        bool RenameCreatesNameConflict(Declaration entity, string newName, out List<Declaration> conflicts);
        bool NewDeclarationCreatesNameConflict(string identifier, DeclarationType declarationType, Declaration parentDeclaration, out List<Declaration> conflicts, Accessibility accessibility = Accessibility.Private);
        bool MoveCreatesNameConflict(Declaration entity, string targetModuleName, Accessibility targetAccessibility, out List<Declaration> conflict, string movedName = null);
    }

    public class NameConflictFinder : INameConflictFinder
    {
        private enum RefactoringAction
        {
            Rename,
            Move,
            New
        };


        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        private delegate bool ConflictFinder(DeclarationProxy proxy, RefactoringAction actionType, out List<Declaration> conflicts);

        private Dictionary<DeclarationType, ConflictFinder> ConflictFinders;

        public NameConflictFinder(IDeclarationFinderProvider declarationFinderProvider)
        {
            _declarationFinderProvider = declarationFinderProvider;

            ConflictFinders = new Dictionary<DeclarationType, ConflictFinder>()
            {
                [DeclarationType.Project] = TryFindProjectNameConflict,
                [DeclarationType.ProceduralModule] = TryFindModuleNameConflict,
                [DeclarationType.ClassModule] = TryFindModuleNameConflict,
                [DeclarationType.Function] = TryFindMemberNameConflict,
                [DeclarationType.Procedure] = TryFindMemberNameConflict,
                [DeclarationType.Property] = TryFindPropertyTypeConflicts,
                [DeclarationType.PropertyGet] = TryFindPropertyNameConflict,
                [DeclarationType.PropertySet] = TryFindPropertyNameConflict,
                [DeclarationType.PropertyLet] = TryFindPropertyNameConflict,
                [DeclarationType.Variable] = TryFindNonMemberNameConflict,
                [DeclarationType.Constant] = TryFindNonMemberNameConflict,
                [DeclarationType.Event] = TryFindEventNameConflict,
                [DeclarationType.Parameter] = TryFindParameterNameConflict,
                [DeclarationType.UserDefinedType] = TryFindUDTNameConflict,
                [DeclarationType.UserDefinedTypeMember] = TryFindUDTMemberNameConflict,
                [DeclarationType.Enumeration] = TryFindEnumerationNameConflict,
                [DeclarationType.EnumerationMember] = TryFindEnumerationMemberNameConflict,
            };
        }

        private static string IncrementIdentifier(string identifier)
        {
            var numeric = string.Concat(identifier.Reverse().TakeWhile(c => char.IsDigit(c)).Reverse());
            if (!int.TryParse(numeric, out var currentNum))
            {
                currentNum = 0;
            }
            var identifierSansNumericSuffix = identifier.Substring(0, identifier.Length - numeric.Length);
            return $"{identifierSansNumericSuffix}{++currentNum}";
        }


        public bool RenameCreatesNameConflict(Declaration entity, string newName, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();
            if (entity.IdentifierName.IsEquivalentVBAIdentifierTo(newName))
            {
                return false;
            }

            var proxy = new DeclarationProxy(entity)
            {
                IdentifierName = newName
            };

            if (!IsPotentialProjectNameConflictType(entity.DeclarationType)
                && !IdentifierIsUsedElsewhereInProject(proxy, newName))
            {
                    return false;
            }

            return EvaluateOperationNameConflict(proxy, RefactoringAction.Rename, out conflicts);
        }

        public bool NewDeclarationCreatesNameConflict(string identifier, DeclarationType declarationType, Declaration parentDeclaration, out List<Declaration> conflicts, Accessibility accessibility = Accessibility.Private)
        {
            conflicts = new List<Declaration>();
            if (!IdentifierIsUsedElsewhereInProject(identifier, parentDeclaration.ProjectId))
            {
                return false;
            }

            var proxy = new DeclarationProxy(identifier, declarationType, parentDeclaration)
            {
                HasPrivateAccessibility = accessibility == Accessibility.Private
            };


            return EvaluateOperationNameConflict(proxy, RefactoringAction.New, out conflicts);
        }

        public bool MoveCreatesNameConflict(Declaration entity, string targetModuleName, Accessibility targetAccessibility, out List<Declaration> conflicts, string movedName = null)
        {
            conflicts = new List<Declaration>();

            var targetModule = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Module)
                .Where(mod => mod.IdentifierName.IsEquivalentVBAIdentifierTo(targetModuleName) && mod.ProjectId == entity.ProjectId).Single();

            var proxy = new DeclarationProxy(entity)
            {
                IdentifierName = movedName ?? entity.IdentifierName,
                HasPrivateAccessibility = targetAccessibility == Accessibility.Private,
                TargetModuleIdentifier = targetModuleName
            };

            if (!proxy.DeclarationType.Equals(DeclarationType.Enumeration) && !IdentifierIsUsedElsewhereInProject(proxy, proxy.IdentifierName))
            {
                return false;
            }

            return EvaluateOperationNameConflict(proxy, RefactoringAction.Move, out conflicts);
        }

        private bool EvaluateOperationNameConflict(DeclarationProxy proxy, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();

            if (ConflictFinders.TryGetValue(proxy.DeclarationType, out var finder))
            {
                return finder(proxy, namingType, out conflicts);
            }
            return true;
        }

        private bool TryFindPropertyTypeConflicts(DeclarationProxy proxy, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();
            var propertyTypes = new DeclarationType[] 
            {
                DeclarationType.PropertyGet,
                DeclarationType.PropertyLet,
                DeclarationType.PropertySet
            };

            foreach (var propertyType in propertyTypes)
            {
                proxy.DeclarationType = propertyType;
                if (EvaluateOperationNameConflict(proxy, namingType, out var propertyConflicts))
                {
                    conflicts.AddRange(propertyConflicts);
                }
            }
            return conflicts.Any();
        }

        //MS-VBAL 5.3.1.6
        //Each subroutine and function must have a procedure name that is different from 
        //any other module variable name, module constant name, enum member name, 
        //or procedure name that is defined within the same module.

        //MS-VBAL 5.3.1.7
        //Each property declaration must have a procedure name that is different from the 
        //name of any other module variable, module constant, enum member name, 
        //external procedure, function, or subroutine that is defined within the same module.
        private bool TryFindMemberNameConflict(DeclarationProxy memberProxy, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();
            if (!IsExistingTargetModule(memberProxy, out var targetModule))
            {
                return false;
            }

            (IEnumerable<Declaration> targetMatches, IEnumerable<Declaration> allMatches) = RelevantIdentifierMatches(memberProxy);

            if (TryFindMemberConflictChecksCommon(memberProxy, targetMatches, out var commonChecks))
            {
                conflicts.AddRange(commonChecks);
            }

            if (NonModuleQualifiedMemberReferenceConflicts(memberProxy, allMatches, namingType, out var refConflicts))
            {
                conflicts.AddRange(refConflicts);
            }
            return conflicts.Any();
        }

        //MS-VBAL 5.3.1.7
        //Each property declaration must have a procedure name that is different from the 
        //name of any other module variable, module constant, enum member name, 
        //external procedure, function, or subroutine that is defined within the same module.
        private bool TryFindPropertyNameConflict(DeclarationProxy memberProxy, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();

            if (!IsExistingTargetModule(memberProxy, out var targetModule))
            {
                return false;
            }

            (IEnumerable<Declaration> targetMatches, IEnumerable<Declaration> allMatches) = RelevantIdentifierMatches(memberProxy);

            var inTargetPropertyIdentifierMatches = targetMatches.Where(idm => idm.DeclarationType.HasFlag(DeclarationType.Property));
            if (inTargetPropertyIdentifierMatches.Any())
            {
                if (!memberProxy.IsExistingDeclaration(out var proxyDeclaration))
                {
                    conflicts.AddRange(inTargetPropertyIdentifierMatches);
                    return true;
                }

                //5.3.1.7 
                //Each property Get must have a unique name
                //Each property Let must have a unique name
                //Each property Set must have a unique name
                if (inTargetPropertyIdentifierMatches.Any(p => p.DeclarationType == memberProxy.DeclarationType))
                {
                    conflicts.AddRange(inTargetPropertyIdentifierMatches.Where(p => p.DeclarationType == memberProxy.DeclarationType));
                    return true;
                }

                //5.3.1.7 each property that shares a common name must have equivalent parameter lists
                if (!HaveEquivalentParameterLists(proxyDeclaration, inTargetPropertyIdentifierMatches.First()))
                {
                    conflicts.AddRange(inTargetPropertyIdentifierMatches);
                }
            }

            if (TryFindMemberConflictChecksCommon(memberProxy, targetMatches.Except(inTargetPropertyIdentifierMatches), out var commonChecks))
            {
                conflicts.AddRange(commonChecks);
            }

            if (NonModuleQualifiedMemberReferenceConflicts(memberProxy, allMatches, namingType, out var nonModuleQualifiedRefDeclarations))
            {
                conflicts.AddRange(nonModuleQualifiedRefDeclarations);
            }
            return conflicts.Any();
        }

        private bool TryFindMemberConflictChecksCommon(DeclarationProxy memberProxy, IEnumerable<Declaration> identifierMatchesInTargetModule, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();
            if (!identifierMatchesInTargetModule.Any()) { return false; }

            if (ModuleLevelElementChecks(identifierMatchesInTargetModule, out var moduleLevelConflicts))
            {
                conflicts.AddRange(moduleLevelConflicts);
            }

            if (LocalDeclarationsHaveSameNameAsParentScope(memberProxy, identifierMatchesInTargetModule, out var localConflicts))
            {
                conflicts.AddRange(localConflicts);
            }

            if (memberProxy.DeclarationType.HasFlag(DeclarationType.Function) 
                    && TryFindFunctionNameMatchesParameterConflict(memberProxy, identifierMatchesInTargetModule, 0, out var parameterConflicts))
            {
                conflicts.AddRange(parameterConflicts);
            }
            return conflicts.Any();
        }

        private bool NonModuleQualifiedMemberReferenceConflicts(DeclarationProxy memberProxy, IEnumerable<Declaration> identifierMatchesAllModules, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();
            if (!IsExistingTargetModule(memberProxy, out var targetModule))
            {
                return false;
            }
            var identifierMatchesExternal = identifierMatchesAllModules.Where(d => d.QualifiedModuleName != targetModule.QualifiedModuleName);

            var nonQualifiedMemberReferences = memberProxy.References.Where(rf => !UsesQualifiedAccess(rf.Context.Parent));

            //NonQualifiedExternalMemberReferences match containing procedure identifier
            conflicts = nonQualifiedMemberReferences.Where(rf => identifierMatchesExternal.Contains(rf.ParentScoping))
                .Select(rf => rf.Declaration)
                .ToList();

            //NonQualifiedExternalMemberReferences match Field or ModuleConstant declaration identifier
            var qmnsWithMatchingFieldOrModuleConstant = identifierMatchesExternal.Where(idm => IsField(idm) || IsModuleConstant(idm)).Select(idm => idm.QualifiedModuleName);
            var matchingRefs = nonQualifiedMemberReferences.Where(rf => qmnsWithMatchingFieldOrModuleConstant.Contains(rf.QualifiedModuleName));

            conflicts.AddRange(matchingRefs.Select(mr => mr.Declaration));

            //NonQualifiedExternalMemberReferences match other nonQualifiedReference(s) within a procedure
            var parentScopeGroupsNonQualifiedMemberReferences = nonQualifiedMemberReferences.GroupBy(rf => rf.ParentScoping);
            foreach (var scopedReferencesGroup in parentScopeGroupsNonQualifiedMemberReferences)
            {
                var matchingNonQualifiedReferences = identifierMatchesAllModules
                                                            .SelectMany(d => d.References)
                                                            .Where(rf => rf.QualifiedModuleName == scopedReferencesGroup.Key.QualifiedModuleName
                                                                                && rf.ParentScoping.Equals(scopedReferencesGroup.Key)
                                                                                && !UsesQualifiedAccess(rf.Context.Parent));

                conflicts.AddRange(matchingNonQualifiedReferences.Select(mnqr => mnqr.Declaration));
            }
            return conflicts.Any();
        }

        private bool TryFindNonMemberNameConflict(DeclarationProxy proxy, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();

            if (!IsExistingTargetModule(proxy, out var targetModule))
            {
                return false;
            }

            (IEnumerable<Declaration> targetMatches, IEnumerable<Declaration> allMatches) = RelevantIdentifierMatches(proxy);

            //Is Module Variable/Constant
            if (proxy.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
            {
                if (ModuleLevelElementChecks(targetMatches, out var moduleLevelConflicts))
                {
                    conflicts.AddRange(moduleLevelConflicts);
                }

                var parentScopedGroups = proxy.References.GroupBy(rf => rf.ParentScoping);
                foreach (var scopedReferencesGroup in parentScopedGroups)
                {
                    if (!scopedReferencesGroup.Key.IsMember())
                    {
                        continue;
                    }

                    var localRefs = _declarationFinderProvider.DeclarationFinder.IdentifierReferences(scopedReferencesGroup.Key.QualifiedName)
                        .Where(lrf => targetModule.QualifiedModuleName == lrf.QualifiedModuleName);

                    var refConflicts = localRefs.Where(lrf => lrf.IdentifierName.IsEquivalentVBAIdentifierTo(proxy.IdentifierName));
                    conflicts.AddRange(refConflicts.Select(rf => rf.Declaration));
                }

                //The nonMember will have Public Accessiblity after the move
                if (namingType == RefactoringAction.Move
                    && (proxy.Template?.HasPrivateAccessibility() ?? !proxy.HasPrivateAccessibility) != proxy.HasPrivateAccessibility
                    && !proxy.HasPrivateAccessibility
                    && allMatches.Any(id => !id.HasPrivateAccessibility()))
                {
                    var publicScopeMatches = allMatches
                                        .Where(match => !match.HasPrivateAccessibility()
                                                    && !AllReferencesUseModuleQualification(match, match.QualifiedModuleName));

                    conflicts.AddRange(publicScopeMatches);
                }
                return conflicts.Any();
            }

            //Is local Variable/Constant
            var idRefs = _declarationFinderProvider.DeclarationFinder.IdentifierReferences(proxy.ParentDeclaration.QualifiedName);

            var idRefConflicts = idRefs.Where(rf => rf.IdentifierName.IsEquivalentVBAIdentifierTo(proxy.IdentifierName));
            conflicts.AddRange(idRefConflicts.Select(rf => rf.Declaration));

            var localConflicts = targetMatches.Where(idm => idm.ParentScopeDeclaration.Equals(proxy.ParentDeclaration)
                                                    && (idm.IsLocalVariable() || idm.IsLocalConstant()));
            conflicts.AddRange(localConflicts);
            return conflicts.Any();
        }

        private bool TryFindProjectNameConflict(DeclarationProxy project, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();

            var publicEnumIdentifierMatchesAllModules = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Enumeration)
                .Where(d => !d.HasPrivateAccessibility() && d.IdentifierName.IsEquivalentVBAIdentifierTo(project.IdentifierName));

            var publicUDTIdentifierMatchesAllModules = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.UserDefinedType)
                    .Where(d => !d.HasPrivateAccessibility() && d.IdentifierName.IsEquivalentVBAIdentifierTo(project.IdentifierName));

            conflicts.AddRange(publicEnumIdentifierMatchesAllModules.Concat(publicUDTIdentifierMatchesAllModules));

            return conflicts.Any();
        }

        private bool TryFindModuleNameConflict(DeclarationProxy moduleProxy, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();
            if (moduleProxy.IsExistingDeclaration(out var moduleDeclaration))
            {
                if (moduleDeclaration.QualifiedModuleName.ProjectName.IsEquivalentVBAIdentifierTo(moduleProxy.IdentifierName))
                {
                    conflicts.Add(moduleProxy.ParentDeclaration);
                }
            }

            (IEnumerable<Declaration> targetMatches, IEnumerable<Declaration> allMatches) = RelevantIdentifierMatches(moduleProxy);

            conflicts.AddRange(allMatches.Where(d => d.DeclarationType.HasFlag(DeclarationType.Module)));
            return conflicts.Any();
        }

        private bool TryFindEventNameConflict(DeclarationProxy eventProxy, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            var identifierMatchesAllModules = _declarationFinderProvider.DeclarationFinder.MatchName(eventProxy.IdentifierName)
                            .Where(match => match.ProjectId == eventProxy.ProjectId).ToList();

            conflicts =  identifierMatchesAllModules.Where(d => d.DeclarationType.HasFlag(DeclarationType.Event)).ToList();
            return conflicts.Any();
        }

        //Ms-VBAL 5.3.1.5
        //1. The name of each <positional-param>, <optional-param>, and <param-array> that are elements of a 
        //function declaration must be different from the name of the function declaration.
        //2. Each <positional-param>, <optional-param>, and <param-array> that are elements of the 
        //same <parameter-list>, <property-parameters>, or <event-parameter-list> must have a distinct names. 
        //3. The name value of a<positional-param>, <optional-param>, or a<param-array> may not be the same 
        //as the name of any variable defined by a <dim-statement>, a<redim-statement>, 
        //or a <const-statement> within the<procedure-body> of the containing procedure declaration.
        private bool TryFindParameterNameConflict(DeclarationProxy parameterProxy, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();

            //1 Different than function
            if (parameterProxy.ParentDeclaration.DeclarationType.HasFlag(DeclarationType.Function)
                && parameterProxy.IdentifierName.IsEquivalentVBAIdentifierTo(parameterProxy.ParentDeclaration.IdentifierName))
            {
                conflicts.Add(parameterProxy.ParentDeclaration);
            }

            //2 Unique params
            var memberScopeMatches = _declarationFinderProvider.DeclarationFinder.MatchName(parameterProxy.IdentifierName)
                                                                        .Where(d => d.ParentScopeDeclaration.IsMember() 
                                                                                && d.QualifiedModuleName.ComponentName == parameterProxy.TargetModuleIdentifier
                                                                                && d.ProjectId == (parameterProxy.Template?.ProjectId ?? string.Empty));
            if (memberScopeMatches.Any())
            {
                conflicts.AddRange(memberScopeMatches);
            }

            //3 Different than existing procedure body references.
            //Strictly speaking, this exceeds 5.3.1.5 (#3 above).  However, changing a parameter
            //to match any referenced element within the procedure body will generate either uncompilable
            //code or change the resulting logic of the procedure.  Flag it as a conflict.
            var procedureBodyReferences = _declarationFinderProvider.DeclarationFinder.IdentifierReferences(parameterProxy.ParentDeclaration.QualifiedName)
                    .Where(rf => rf.IdentifierName.IsEquivalentVBAIdentifierTo(parameterProxy.IdentifierName));
            if (procedureBodyReferences.Any())
            {
                conflicts.AddRange(procedureBodyReferences.Select(rf => rf.Declaration));
            }

            return conflicts.Any();
        }

        //MS-VBAL 5.2.3.3 UserDefinedType Declarations
        private bool TryFindUDTNameConflict(DeclarationProxy udtProxy, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();

            //UserDefinedType and Enumerations have the same conflict rules
            if (UdtAndEnumerationConflicts(udtProxy, namingType, out var udtOrEnumConflicts))
            {
                conflicts.AddRange(udtOrEnumConflicts);
            }

            return conflicts.Any();
        }

        private bool TryFindUDTMemberNameConflict(DeclarationProxy udtMemberProxy, RefactoringAction naming, out List<Declaration> conflicts)
        {
            var relatedMembers = _declarationFinderProvider.DeclarationFinder.Members(udtMemberProxy.ParentDeclaration.QualifiedModuleName)
                                                .Where(d => d.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember) && d.ParentDeclaration == udtMemberProxy.ParentDeclaration);
            conflicts = relatedMembers.Where(rm => rm.IdentifierName.IsEquivalentVBAIdentifierTo(udtMemberProxy.IdentifierName)).ToList();
            return conflicts.Any();
        }

        private bool TryFindAllMatchesForProposedIdentifier(string projectID, string proposedIdentifier, out IEnumerable<Declaration> matches)
        {
            matches = _declarationFinderProvider.DeclarationFinder.MatchName(proposedIdentifier)
                                        .Where(match => match.ProjectId == projectID).ToList();

            return matches.Any();
        }

        //MS-VBAL 5.2.3.4 Enum Declarations
        private bool TryFindEnumerationNameConflict(DeclarationProxy enumerationProxy, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();

            var destinationModuleDeclarations = GetTargetModuleMembers(enumerationProxy);
            if (!destinationModuleDeclarations.Any())
            {
                conflicts = new List<Declaration>();
                return false;
            }

            //UserDefinedType and Enumerations have the same conflit rules
            if (UdtAndEnumerationConflicts(enumerationProxy, namingType, out var udtOrEnumConflicts))
            {
                conflicts.AddRange(udtOrEnumConflicts);
            }

            if (namingType == RefactoringAction.Move)
            {
                var enumMembers = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.EnumerationMember)
                        .Where(en => en.ParentDeclaration == enumerationProxy.Template);

                foreach (var enumMember in enumMembers)
                {
                    var enumMemberProxy = new DeclarationProxy(enumMember)
                    {
                        TargetModuleIdentifier = enumerationProxy.TargetModuleIdentifier
                    };

                    if (TryFindEnumerationMemberNameConflict(enumMemberProxy, namingType, out var memberConflicts))
                    {
                        conflicts.AddRange(memberConflicts);
                    }
                }
            }

            return conflicts.Any();
        }

        private bool TryFindEnumerationMemberNameConflict(DeclarationProxy enumerationMemberProxy, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            var destinationModuleDeclarations = GetTargetModuleMembers(enumerationMemberProxy);
            if (!destinationModuleDeclarations.Any())
            {
                conflicts = new List<Declaration>();
                return false;
            }

            var identifierMatchingDeclarations
                            = destinationModuleDeclarations.Where(d => d != enumerationMemberProxy.Template 
                                        && d.IdentifierName.IsEquivalentVBAIdentifierTo(enumerationMemberProxy.IdentifierName));

            return ModuleLevelElementChecks(identifierMatchingDeclarations, out conflicts);
        }

        //MS-VBAL 5.3.1.6
        //Each subroutine and function must have a procedure name that is different from 
        //any other module variable name, module constant name, enum member name, 
        //or procedure name that is defined within the same module.

        //MS-VBAL 5.3.1.7
        //Each property declaration must have a procedure name that is different from the 
        //name of any other module variable, module constant, enum member name, 
        //external procedure, function, 
        //or subroutine that is defined within the same module.
        private bool ModuleLevelElementChecks(IEnumerable<Declaration> matchingDeclarations, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();
            foreach (var identifierMatch in matchingDeclarations)
            {
                if (identifierMatch.IsMember()
                    || IsField(identifierMatch)
                    || IsModuleConstant(identifierMatch)
                    || IsEnumMember(identifierMatch))
                {
                    conflicts.Add(identifierMatch);
                }
            }
            return conflicts.Any();
        }

        private bool LocalDeclarationsHaveSameNameAsParentScope(DeclarationProxy proxy, IEnumerable<Declaration> identifierMatchesInTargetModule, out List<Declaration> conflicts)
        {
            //MS-VBAL 5.4.3.1, 5.4.3.2
            //A local variable/constant cannot have the same name as the containing procedure name.
            conflicts =  identifierMatchesInTargetModule.Where(idm => (idm.IsLocalVariable()
                                                        || idm.IsLocalConstant())
                                                        && idm.ParentScopeDeclaration.Equals(proxy.Template)).ToList();
            return conflicts.Any();
        }

        private bool TryFindFunctionNameMatchesParameterConflict(DeclarationProxy memberProxy, IEnumerable<Declaration> identifierMatchesAllModules, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();
            var parameters = identifierMatchesAllModules
                                    .Where(d => memberProxy.DeclarationType.HasFlag(DeclarationType.Function)
                                                    && d.DeclarationType.HasFlag(DeclarationType.Parameter)
                                                    && (memberProxy.Template?.Equals(d.ParentScopeDeclaration) ?? false) );
            conflicts.AddRange(parameters); 
            return conflicts.Any();
        }

        private bool UdtAndEnumerationConflicts(DeclarationProxy proxyEntity, RefactoringAction namingType, out List<Declaration> conflicts)
        {
            conflicts = new List<Declaration>();
            var destinationModuleDeclarations = GetTargetModuleMembers(proxyEntity);

            var udtIdentifierConflictTypes = new List<DeclarationType>()
            {
                DeclarationType.UserDefinedType,
                DeclarationType.Enumeration,
            };

            foreach (var potentialConflict in destinationModuleDeclarations.Where(pc => pc.IdentifierName.IsEquivalentVBAIdentifierTo(proxyEntity.IdentifierName)))
            {
                if (udtIdentifierConflictTypes.Any(ect => potentialConflict.DeclarationType.HasFlag(ect)))
                {
                    conflicts.Add(potentialConflict);
                }
            }

            if (!proxyEntity.HasPrivateAccessibility)
            {
                var conflictingModuleIdentifiers = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Module)
                    .Where(m => m.ProjectId == proxyEntity.ProjectId
                                                    && m.IdentifierName.IsEquivalentVBAIdentifierTo(proxyEntity.IdentifierName));
                conflicts.AddRange(conflictingModuleIdentifiers);

                var conflictingProjectIdentifiers = _declarationFinderProvider.DeclarationFinder.Projects
                    .Where(p => p.IdentifierName.IsEquivalentVBAIdentifierTo(proxyEntity.IdentifierName));

                conflicts.AddRange(conflictingProjectIdentifiers);

                var conflictingUDTsInProject = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.UserDefinedType)
                        .Where(udtCandidate => udtCandidate.ProjectId == proxyEntity.ProjectId
                                                    && udtCandidate != proxyEntity.Template
                                                    && !udtCandidate.HasPrivateAccessibility()
                                                    && udtCandidate.IdentifierName.IsEquivalentVBAIdentifierTo(proxyEntity.IdentifierName));
                conflicts.AddRange(conflictingUDTsInProject);

                var conflictingEnumsInProject = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Enumeration)
                        .Where(enumCandidate => enumCandidate.ProjectId == proxyEntity.ProjectId
                                                    && enumCandidate != proxyEntity.Template
                                                    && !enumCandidate.HasPrivateAccessibility()
                                                    && enumCandidate.IdentifierName.IsEquivalentVBAIdentifierTo(proxyEntity.IdentifierName));
                conflicts.AddRange(conflictingEnumsInProject);
            }
            return conflicts.Any();
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

        private bool IdentifierIsUsedElsewhereInProject(IDeclarationProxy proxy, string newName)
        {
            if (proxy.IsExistingDeclaration(out var declaration))
            {
                return _declarationFinderProvider.DeclarationFinder.MatchName(newName)
                                .Any(matchedName => matchedName != declaration && matchedName.ProjectId == proxy.ProjectId);
            }
            return _declarationFinderProvider.DeclarationFinder.MatchName(newName)
                            .Any(matchedName => matchedName.ProjectId == proxy.ProjectId);
        }

        private bool IdentifierIsUsedElsewhereInProject(string identifier, string projectID)
        {
            return _declarationFinderProvider.DeclarationFinder.MatchName(identifier)
                            .Any(matchedName => matchedName.ProjectId == projectID);
        }

        private bool IsField(Declaration declaration)
        {
            return declaration.IsVariable() && !declaration.IsLocalVariable();
        }

        private bool IsModuleConstant(Declaration declaration)
        {
            return declaration.IsConstant() && !declaration.IsLocalConstant();
        }

        private bool IsEnumMember(Declaration declaration)
        {
            return declaration.DeclarationType.Equals(DeclarationType.EnumerationMember);
        }

        private bool IsPotentialProjectNameConflictType(DeclarationType declarationType)
        {
            return declarationType.HasFlag(DeclarationType.Enumeration)
                || declarationType.HasFlag(DeclarationType.UserDefinedType)
                || declarationType.HasFlag(DeclarationType.Project);
        }

        private bool UsesQualifiedAccess(RuleContext ruleContext)
        {
            return (ruleContext is VBAParser.WithMemberAccessExprContext)
                || (ruleContext is VBAParser.MemberAccessExprContext);
        }

        private IEnumerable<Declaration> GetTargetModuleMembers(IDeclarationProxy proxy)
        {
            var modules = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Module)
                .Where(mod => mod.ProjectId == (proxy.Template?.ProjectId ?? string.Empty) && mod.IdentifierName == proxy.TargetModuleIdentifier);
            if (modules.Any())
            {
                return _declarationFinderProvider.DeclarationFinder.Members(modules.Single().QualifiedModuleName);
            }
            return Enumerable.Empty<Declaration>();
        }

        private bool IsExistingTargetModule(IDeclarationProxy proxy, out Declaration targetModule)
        {
            targetModule = null;
            var modules = _declarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Module)
                .Where(mod => mod.ProjectId == proxy.ProjectId && mod.IdentifierName == proxy.TargetModuleIdentifier);
            if (modules.Any())
            {
                targetModule = modules.Single();
                return true;
            }
            return false;
        }

        private (IEnumerable<Declaration> targetModuleMatches, IEnumerable<Declaration> allModuleMatches) RelevantIdentifierMatches(IDeclarationProxy proxy)
        {
            var matchesAllModules = _declarationFinderProvider.DeclarationFinder.MatchName(proxy.IdentifierName)
                            .Where(match => match.ProjectId == proxy.ProjectId && match != proxy.Template);

            var targetModuleMatches = matchesAllModules.Where(mod => mod.QualifiedModuleName.ComponentName == proxy.TargetModuleIdentifier)
                            .Where(match => match.ProjectId == proxy.ProjectId && match != proxy.Template);

            return (targetModuleMatches, matchesAllModules);
        }

        private bool AllReferencesUseModuleQualification(Declaration declaration, QualifiedModuleName declarationQMN)
        {
            var referencesToModuleQualify = declaration.References.Where(rf => (rf.QualifiedModuleName != declarationQMN));

            return referencesToModuleQualify.All(rf => UsesQualifiedAccess(rf.Context.Parent));
        }

        public interface IDeclarationProxy
        {
            Declaration Template { get; }
            string ProjectId { get; }
            string IdentifierName { set;  get; }
            Declaration ParentDeclaration { set; get; }
            DeclarationType DeclarationType { set; get; }
            Accessibility Accessibility { set; get; }
            string TargetModuleIdentifier { set; get; }
            bool IsExistingDeclaration(out Declaration declaration);
        }

        private class DeclarationProxy : IDeclarationProxy
        {
            private readonly Declaration _declaration;
            public DeclarationProxy(Declaration declaration)
                : this(declaration.IdentifierName, declaration.DeclarationType, declaration.ParentDeclaration)
            {
                _declaration = declaration;
                TargetModuleIdentifier = declaration.ComponentName;
                HasPrivateAccessibility = _declaration.Accessibility.Equals(Accessibility.Private);
            }

            public DeclarationProxy(string identifier, DeclarationType declarationType, Declaration parentDeclaration, Accessibility accessibility = Accessibility.Private)
            {
                ParentDeclaration = parentDeclaration;
                IdentifierName = identifier;
                DeclarationType = declarationType;
                HasPrivateAccessibility = accessibility == Accessibility.Private;
            }

            public Declaration Template => _declaration;
            public string IdentifierName { set;  get; }
            public Declaration ParentDeclaration { set; get; }
            public DeclarationType DeclarationType { set; get; }
            public bool HasPrivateAccessibility { set; get; }
            public IEnumerable<IdentifierReference> References => _declaration?.References ?? Enumerable.Empty<IdentifierReference>();
            public string ProjectId => _declaration?.ProjectId ?? ParentDeclaration?.ProjectId ?? string.Empty;
            public Accessibility Accessibility { set; get; }
            public string TargetModuleIdentifier { set; get; }
            public bool IsExistingDeclaration(out Declaration declaration)
            {
                declaration = _declaration;
                return _declaration != null;
            }
         }
    }
}

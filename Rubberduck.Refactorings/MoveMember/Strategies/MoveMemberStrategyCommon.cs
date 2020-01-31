using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberRefactoringStrategy
    {
        void ModifyContent();
        string DestinationNewModuleContent { get; }
        int DestinationNewContentLineCount { get; }
        string PreviewDestination();
        //Testing only?
        string DestinationMemberCodeBlock(Declaration member);
    }

    public class MoveMemberStrategyCommon
    {
        private IMoveScenario _scenario;
        private IProvideMoveDeclarationGroups _moveDeclarationGroups;
        private MoveMemberRewritingManager _rewritingManager;
        private MoveMemberContentInfo _moveMemberContentInfo;

        private static List<DeclarationType> NeverAddMemberAccessTypes = new List<DeclarationType>()
        {
            DeclarationType.UserDefinedType,
            DeclarationType.UserDefinedTypeMember,
            DeclarationType.Enumeration,
            DeclarationType.EnumerationMember
        };

        public MoveMemberStrategyCommon(IMoveScenario scenario, MoveMemberRewritingManager rewritingManager)
        {
            _scenario = scenario;
            _rewritingManager = rewritingManager;
            _moveDeclarationGroups = scenario as IProvideMoveDeclarationGroups;
            _moveMemberContentInfo = new MoveMemberContentInfo(_moveDeclarationGroups);

            ModifySourceModuleAction = null;
            UpdateSourceReferencesToMovedElementsAction = UpdateSourceReferencesToMovedElements;
            DestinationMemberCodeBlockFunc = DestinationMemberCodeBlockDefault;
            DestinationNonMemberCodeBlockFunc = DestinationNonMemberCodeBlockDefault;
            DestinationMemberCodeBlockRewriterFunc = DestinationMemberCodeBlockRewriter;
            PrepareNewDestinationCodeElementsAction = PrepareNewDestinationCodeElementsDefault;
        }

        public Action<IMoveEndpointRewriter> ModifySourceModuleAction { set; get; }

        public Action<IMoveEndpointRewriter> PrepareNewDestinationCodeElementsAction { set; get; } = null;

        public Func<Declaration, string> DestinationMemberCodeBlockFunc { set; get; }

        public Func<Declaration, IMoveEndpointRewriter, IMoveEndpointRewriter> DestinationMemberCodeBlockRewriterFunc { set; get; }

        public Func<Declaration, string> DestinationNonMemberCodeBlockFunc { set; get; }

        public Action<IMoveEndpointRewriter> UpdateSourceReferencesToMovedElementsAction { set; get; }

        public string PreviewDestination()
        {
            var preview = string.Empty;
            if (_scenario.CreatesNewModule)
            {
                var tempRewriter = _rewritingManager.CheckOutTemporaryRewriter(_scenario.SourceContentProvider.QualifiedModuleName);
                PrepareNewDestinationCodeElementsAction(tempRewriter);
                preview = $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}{_scenario.DestinationContentProvider.AllNewContent}";
            }
            else
            {
                var destinationRewriter = _rewritingManager.CheckOutTemporaryRewriter(_scenario.DestinationContentProvider.QualifiedModuleName);
                ModifyDestination(destinationRewriter);
                preview = destinationRewriter.GetText();
            }

            _scenario.DestinationContentProvider.ResetContent();
            return preview;
        }

        public void UpdateSourceReferencesToMovedElements(IMoveEndpointRewriter rewriter)
        {
            //var referencesToMovedElements = _moveDeclarationGroups.Moving.AllDeclarations.Where(d => !_moveDeclarationGroups.IsPropertyBackingVariable(d)).AllReferences()
            //    .Where(rf => _moveDeclarationGroups.Retain.AllDeclarations
            //                .Except(_moveDeclarationGroups.Forward).Contains(rf.ParentScoping)
            //                    && !_moveDeclarationGroups.Forward.Contains(rf.Declaration));


            var sourceIdentifierReferencesToModuleQualify = RenameSupport.RenameReferencesByQualifiedModuleName(_moveDeclarationGroups.SelectedElements.AllReferences())
                .Where(g => g.Key == _scenario.QualifiedModuleNameSource)
                .SelectMany(grp => grp)
                .Where(rf => !_moveDeclarationGroups.Moving.Contains(rf.ParentScoping));

            foreach (var rf in sourceIdentifierReferencesToModuleQualify)
            {
                rewriter.Replace(rf.Context, ScopeResolvedName(rf));
            }
        }

        public static bool IsSingleDeclarationSelection(IProvideMoveDeclarationGroups groups, DeclarationType declarationType)
        {
            if (groups is null || !groups.SelectedElements.AllDeclarations.Any()) { return false; }

            if (declarationType.HasFlag(DeclarationType.Property))
            {
                return groups.IsSingleDeclarationSelection;
            }

            return groups.IsSingleDeclarationSelection
                && (groups.SelectedElements.FirstOrDefault()?.DeclarationType.Equals(declarationType) ?? false);
        }

        public static bool IsUnsupportedMoveGeneral(IMoveScenario scenario, IProvideMoveDeclarationGroups groups)
        {
            if (scenario.MoveDefinition.Endpoints.Equals(MoveEndpoints.Undefined)) { return true; }

            if (scenario.MoveDefinition.IsClassModuleDestination) { return true; }

            if (scenario.MoveDefinition.Destination.ComponentType.Equals(ComponentType.UserForm)) { return true; }

            if (CausesDestinationNameConflicts(scenario.DeclarationFinderProvider, scenario.MoveDefinition, groups)) { return true; }

            return false;
        }

        public static bool IsUnsupportedMoveGeneralMethod(IMoveScenario scenario, IProvideMoveDeclarationGroups groups)
        {
            if (!HasMethodSelection(groups)) { return false; }

            if (MoveDefinitionIncludesLifeCycleHandler(scenario.MoveDefinition, groups)) { return true; }

            if (IsOrReferencesEventSink(groups)) { return true; }

            if (ReferencesOrRaisesAnEvent(groups)) { return true; }

            if (IsUserFormEvent(groups, scenario.MoveDefinition)) { return true; }

            if (IsOrReferencesAnInterfaceImplementation(groups, scenario.MoveDefinition)) { return true; }

            if (IsOrReferencesAnInterfaceDefinition(groups, scenario.MoveDefinition)) { return true; }

            return false;
        }

        public static bool MoveDefinitionIncludesLifeCycleHandler(MoveDefinition moveDefinition, IProvideMoveDeclarationGroups groups)
        {
            if (moveDefinition.IsClassModuleSource || moveDefinition.IsClassModuleDestination)
            {
                var lifecycleHandlers = new List<string>() { MoveMemberResources.Class_Initialize, MoveMemberResources.Class_Terminate };
                if (groups.MoveAndDelete.Concat(groups.Forward).Any(m => lifecycleHandlers.Contains(m.IdentifierName)))
                {
                    return true;
                }
            }
            return false;
        }

        public static bool CausesDestinationNameConflicts(IDeclarationFinderProvider declarationFinderProvider, MoveDefinition moveDefinition, IProvideMoveDeclarationGroups groups)
        {
            if (moveDefinition.Destination.Module is null) { return false; }

            var destinationDeclarations = moveDefinition.Destination.Module != null ?
                declarationFinderProvider.DeclarationFinder.Members(moveDefinition.Destination.Module)
                : Enumerable.Empty<Declaration>();

            var nameConflicts = destinationDeclarations.Where(dec => groups.MoveAndDelete.Concat(groups.Forward).Any(nm => nm.IdentifierName.Equals(dec.IdentifierName)));
            return nameConflicts.Any();
        }


        public static bool IsOrReferencesEventSink(IProvideMoveDeclarationGroups groups)
        {
            var eventSinkPrefixes = groups.MoveableElements.Where(se => se.IsVariable()
                && se.IsWithEvents).Select(we => $"{we.IdentifierName}_");

            if (eventSinkPrefixes.Any())
            {
                var memberNames = groups.Participants.AllDeclarations.Select(m => m.IdentifierName);

                if (memberNames.Any(mn => eventSinkPrefixes.Any(p => mn.StartsWith(p))))
                {
                    return true;
                }
            }
            return false;
        }

        public static bool IsUserFormEvent(IProvideMoveDeclarationGroups groups, MoveDefinition moveDefinition)
        {
            return moveDefinition.IsUserFormSource && groups.MoveAndDelete.Concat(groups.Forward).Any(m => m.IdentifierName.StartsWith($"{MoveMemberResources.UserForm}_"));
        }

        public static bool IsOrReferencesAnInterfaceImplementation(IProvideMoveDeclarationGroups groups, MoveDefinition moveDefinition)
        {
            return (moveDefinition.IsClassModuleSource
                        || moveDefinition.IsUserFormSource)
                && groups.MoveAndDelete.Concat(groups.Forward).Any(m => m is ModuleBodyElementDeclaration mbe && mbe.IsInterfaceImplementation);
        }

        public static bool IsOrReferencesAnInterfaceDefinition(IProvideMoveDeclarationGroups groups, MoveDefinition moveDefinition)
        {
            return (moveDefinition.IsClassModuleSource
                        || moveDefinition.IsUserFormSource)
                && groups.MoveAndDelete.Concat(groups.Forward).Any(m => m is ModuleBodyElementDeclaration mbe && mbe.IsInterfaceMember);
        }

        //TODO: No test exists...probably has to be this way
        public static bool IsAControl(IProvideMoveDeclarationGroups groups, MoveDefinition moveDefinition)
        {
            return moveDefinition.IsUserFormSource && groups.SelectedElements.NonMembers
                .Any(m => m.DeclarationType.HasFlag(DeclarationType.Control));
        }

        public static bool ReferencesOrRaisesAnEvent(IProvideMoveDeclarationGroups groups)
        {
            var eventRefs = groups.AllDeclarations.Where(m => m.DeclarationType.HasFlag(DeclarationType.Event)).AllReferences();

            var referencesToAnEvent = groups.MoveAndDelete.Concat(groups.Forward).Where(d => d.IsMember())
                .Where(m => eventRefs.Any(rf => rf.ParentScoping == m));

            return referencesToAnEvent.Any();
        }

        public string ForwardToModuleLExpression()
        {
            var variablePrefix = _scenario.MoveDefinition.IsClassModuleDestination ?
                $"{MoveMemberResources.Prefix_Variable}"
                : string.Empty;

            var newVariableName = $"{variablePrefix}{_scenario.DestinationContentProvider.ModuleName}";

            var destinationClassInstanceVariables = DestinationClassInstanceVariables;
            if (_scenario.CreatesNewModule)
            {
                return newVariableName;
            }

            switch (destinationClassInstanceVariables.Count())
            {
                case 0:
                    return newVariableName;
                case 1:
                    return destinationClassInstanceVariables.First().IdentifierName;
                default:
                    throw new MoveMemberUnsupportedMoveException(destinationClassInstanceVariables.First());
            }
        }

        public string DestinationNewModuleContent
            => _scenario.DestinationContentProvider.AllNewContent;

        public int DestinationNewContentLineCount
            => _scenario.DestinationContentProvider.NewContentLineCount;

        public void ModifyContent(Action<IMoveEndpointRewriter> modifySourceAction, Action<IMoveEndpointRewriter> modifyDestinationAction = null)
        {
#if DEBUG
            var preModify = GetContentState();
#endif
            ModifySource(modifySourceAction);
#if DEBUG
            var postSourceModify = GetContentState();
#endif
            ModifyDestination();
#if DEBUG
            var postDestinationModify = GetContentState();
#endif
            UpdateCallSites();
#if DEBUG
            var finalState = GetContentState();
#endif
        }

        public static bool IsPubliclyAccessibleInStdModule(Declaration dec) =>
            dec.QualifiedModuleName.ComponentType.Equals(ComponentType.StandardModule) && dec.Accessibility.Equals(Accessibility.Public);

        private static bool HasMethodSelection(IProvideMoveDeclarationGroups groups)
        {
            if (groups is null || !groups.SelectedElements.AllDeclarations.Any()) { return false; }

            return groups.SelectedElements.AllDeclarations.Any(se => se.IsMember());
        }

        private void ModifySource(Action<IMoveEndpointRewriter> modifySourceAction)
        {
            var sourceRewriter = _rewritingManager.CheckOutEndpointRewriter(_scenario.SourceContentProvider.QualifiedModuleName);
            modifySourceAction(sourceRewriter);
        }

        private void ModifyDestination(IMoveEndpointRewriter rewriter = null)
        {
            if (_scenario.CreatesNewModule) { return; }

            var endpointRewriter = rewriter ?? _rewritingManager.CheckOutEndpointRewriter(_scenario.DestinationContentProvider.QualifiedModuleName);
            var movedElementReferencesInDestination = _moveDeclarationGroups.SelectedElements.AllReferences()
                .Where(rf => _scenario.DestinationContentProvider.ContainsReference(rf));

            if (movedElementReferencesInDestination.Any())
            {
                endpointRewriter.RemoveMemberAccess(movedElementReferencesInDestination);

                endpointRewriter.RemoveWithMemberAccess(movedElementReferencesInDestination);
            }

            var tempRewriter = _rewritingManager.CheckOutTemporaryRewriter(_scenario.SourceContentProvider.QualifiedModuleName);
            PrepareNewDestinationCodeElementsAction(tempRewriter);

            endpointRewriter = _scenario.DestinationContentProvider.InsertNewContent(endpointRewriter);
        }

        public void ReplaceMovedOrRenamedReferenceIdentifiers(IMoveEndpointRewriter rewriter)
        {
            foreach (var v in _moveDeclarationGroups.VariableReferenceReplacement)
            {
                var declaration = v.Key;
                foreach (var rf in declaration.References)
                {
                    if (!_moveDeclarationGroups.Participants.AllDeclarations.Contains(rf.ParentScoping))
                    {
                        if (_moveDeclarationGroups.IsPropertyBackingVariable(declaration))
                        {
                            var lifeCycleMember = _moveDeclarationGroups.Retain.Members.FirstOrDefault(m => m.IdentifierName.Equals(MoveMemberResources.Class_Initialize));
                            if (lifeCycleMember != null)
                            {
                                //TODO: this may be ok for classes, but initalizing a value-type needs to be replicated in the destination
                                if (rf.Context.IsDescendentOf(lifeCycleMember.Context)
                                    && rf.Context.Parent is VBAParser.SetStmtContext setStmtCtxt)
                                {
                                    rewriter.Remove(setStmtCtxt);
                                }
                                continue;
                            }
                        }
                        rewriter.Replace(rf.Context, v.Value);
                    }
                }
            }
        }

        public void RemoveDeclarations(IMoveEndpointRewriter rewriter)
            => rewriter.RemoveDeclarations(_moveDeclarationGroups.Remove, _moveDeclarationGroups.AllDeclarations);

        public void UpdateCallSites()
        {
            if (!_scenario.MoveDefinition.IsStdModuleDestination)
            {
                return;
            }

            var qmnToReferenceIdentifiers = RenameSupport.RenameReferencesByQualifiedModuleName(_moveDeclarationGroups.SelectedElements.AllReferences());

            foreach (var references in qmnToReferenceIdentifiers.Where(qmn => !_scenario.IsMoveEndpoint(qmn.Key)))
            {
                var moduleRewriter = _rewritingManager.CheckOutModuleRewriter(references.Key);

                var memberAccessExprContexts = references.Where(rf => rf.Context.Parent is VBAParser.MemberAccessExprContext && rf.Context is VBAParser.UnrestrictedIdentifierContext)
                        .Select(rf => rf.Context.Parent as VBAParser.MemberAccessExprContext);

                foreach (var memberAccessExprContext in memberAccessExprContexts)
                {
                    moduleRewriter.Replace(memberAccessExprContext.lExpression(), _scenario.DestinationContentProvider.ModuleName);
                }

                var withMemberAccessExprContexts = references.Where(rf => rf.Context.Parent is VBAParser.WithMemberAccessExprContext)
                        .Select(rf => rf.Context.Parent as VBAParser.WithMemberAccessExprContext);

                foreach (var wma in withMemberAccessExprContexts)
                {
                    moduleRewriter.InsertBefore(wma.Start.TokenIndex, _scenario.DestinationContentProvider.ModuleName);
                }

                var nonQualifiedReferences = references.Where(rf => !(rf.Context.Parent is VBAParser.WithMemberAccessExprContext
                    || (rf.Context.Parent is VBAParser.MemberAccessExprContext && rf.Context is VBAParser.UnrestrictedIdentifierContext)));

                foreach (var rf in nonQualifiedReferences)
                {
                    moduleRewriter.Replace(rf.Context, $"{_scenario.DestinationContentProvider.ModuleName}.{rf.Declaration.IdentifierName}");
                }
            }
        }

        public void EnsureClassIsValidWhereReferenced(IMoveEndpointRewriter rewriter)
        {
            if (!_scenario.MoveDefinition.IsStdModuleSource)
            {
                return;
            }

            var membersReferencingMovedMembers = _moveDeclarationGroups.MoveAndDelete.NonMembers.Where(d => !_moveDeclarationGroups.IsPropertyBackingVariable(d)).AllReferences()
                .Where(rf => _moveDeclarationGroups.Retain.Members.Contains(rf.ParentScoping))
                .Select(rf => rf.ParentScoping)
                .Distinct();

            var membersRequiringClassVariable =
                _moveDeclarationGroups.Forward.Where(d => d.IsMember())
                    .Concat(membersReferencingMovedMembers).Distinct();
            var instantiationCall = $"{_scenario.SourceContentProvider.ClassInstantiationSubName}{Environment.NewLine}{Environment.NewLine}";

            foreach (var member in membersRequiringClassVariable)
            {
                var blockContext = member.Context.GetDescendent<VBAParser.BlockContext>();
                if (blockContext != null)
                {
                    rewriter.InsertBefore(blockContext.Start.TokenIndex, instantiationCall);
                }
            }
        }

        public void InsertNewSourceContent(IMoveEndpointRewriter rewriter)
        {
            if (DestinationClassInstanceVariables.Count() == 0 && _scenario.MoveDefinition.IsClassModuleSource)
            {
                var classInitialize = _scenario.SourceContentProvider.ModuleDeclarations.FirstOrDefault(el => el.IdentifierName.Equals(MoveMemberResources.Class_Initialize));
                if (classInitialize != null)
                {
                    if (!rewriter.GetText().Contains(ClassInstantiationFragment))
                    {
                        rewriter.InsertBeforeDescendentContext<VBAParser.BlockContext>(classInitialize, $"{ClassInstantiationFragment}{Environment.NewLine}");
                    }
                }
            }

            _scenario.SourceContentProvider.InsertNewContent(rewriter);
        }

        public string ClassInstantiationFragment => $"{_scenario.SourceContentProvider.ClassInstantiationFragment}";

        public string CallSiteArguments(Declaration member)
            => _moveMemberContentInfo.CallSiteArguments(member);

        public IEnumerable<Declaration> DestinationClassInstanceVariables
            => RetrieveDestinationClassVariables(_scenario, _moveDeclarationGroups);

        public static IEnumerable<Declaration> RetrieveDestinationClassVariables(IMoveScenario scenario, IProvideMoveDeclarationGroups groups)
            => groups.AllDeclarations.Where(el => el.IsVariable()
                    && (el.AsTypeDeclaration?.IdentifierName.Equals(scenario.DestinationContentProvider.Module.IdentifierName) ?? false));

        public void PrepareNewDestinationCodeElementsDefault(IMoveEndpointRewriter tempRewriter)
        {
            foreach (var element in _moveDeclarationGroups.Moving.Members)
            {
                var codeBlock = DestinationMemberCodeBlockFunc(element);
                _scenario.DestinationContentProvider.AddCodeBlock(codeBlock);
            }

            foreach (var element in _moveDeclarationGroups.Moving.NonMembers)
            {
                var declarationBlock = DestinationNonMemberCodeBlockFunc(element);
                _scenario.DestinationContentProvider.AddDeclarationBlock(declarationBlock);
            }
        }

        public string DestinationMemberCodeBlockDefault(Declaration member)
        {
            var tempRewriter = _rewritingManager.CheckOutTemporaryRewriter(_scenario.SourceContentProvider.QualifiedModuleName);
            tempRewriter = DestinationMemberCodeBlockRewriterFunc(member, tempRewriter);

            var destinationMemberAccessReferencesToMovedMembers = _scenario.DestinationContentProvider.ModuleDeclarations
                .AllReferences().Where(rf => rf.ParentScoping == member);

            tempRewriter.RemoveMemberAccess(destinationMemberAccessReferencesToMovedMembers);
            return tempRewriter.GetModifiedText(member);
        }

        public bool RequiresForcedPublicVisibility(Declaration member)
            => _moveDeclarationGroups.SelectedElements.Contains(member)
                || _moveDeclarationGroups.Forward.Contains(member)
                || !_scenario.IsOnlyReferencedByMovedElements(member);


        public (string Source, string Destination) GetContentState()
        {
#if DEBUG
            var sourceRewriter = _rewritingManager.CheckOutEndpointRewriter(_scenario.SourceContentProvider.QualifiedModuleName);
            var destinationRewriter = _rewritingManager.CheckOutEndpointRewriter(_scenario.DestinationContentProvider.QualifiedModuleName);

            var newSource = sourceRewriter.GetText();
            var newDestination = destinationRewriter.GetText();
            return (newSource, newDestination);
#else
            return (string.Empty, stringEmpty);
#endif
        }

        //TODO: This needs to be distributed to the various types of strategies - Variable, Procedure, Function, etc
        private string ScopeResolvedName(IdentifierReference identifierReference)
        {
            if (NeverAddMemberAccessTypes.Contains(identifierReference.Declaration.DeclarationType))
            {
                return identifierReference.IdentifierName;
            }

            if (_moveDeclarationGroups.TryGetPropertiesFromBackingVariable(identifierReference.Declaration, out List<Declaration> properties))
            {
                return $"{ForwardToModuleLExpression()}.{properties.First().IdentifierName}";
            }

            if (identifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Variable)
                || identifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Procedure))
            {
                return $"{ForwardToModuleLExpression()}.{identifierReference.IdentifierName}";
            }

            if (identifierReference.Declaration.DeclarationType.HasFlag(DeclarationType.Function)
                && identifierReference.IsAssignment)
            {
                return identifierReference.IdentifierName;
            }

            if (_moveDeclarationGroups.Retain.Contains(identifierReference.Declaration))
            {
                return identifierReference.IdentifierName;
            }
            return $"{_scenario.DestinationContentProvider.ModuleName}.{identifierReference.IdentifierName}";
        }

        private IMoveEndpointRewriter DestinationMemberCodeBlockRewriter(Declaration member, IMoveEndpointRewriter tempRewriter)
        {
            if (RequiresForcedPublicVisibility(member))
            {
                tempRewriter.SetMemberAccessibility(member, Tokens.Public);
            }

            if (_scenario.MoveDefinition.IsStdModuleSource)
            {
                var declarationGroups = _scenario as IProvideMoveDeclarationGroups;
                if (declarationGroups.Retain.PublicNonMembers.Any())
                {
                    foreach (var publicVariable in declarationGroups.Retain.PublicNonMembers)
                    {
                        var destinationRefs = publicVariable.References.Where(rf => declarationGroups.Moving.Contains(rf.ParentScoping));
                        foreach (var rf in destinationRefs)
                        {
                            tempRewriter.Replace(rf.Context, $"{_scenario.MoveDefinition.Source.ModuleName}.{rf.IdentifierName}");
                        }
                    }
                }

                if (declarationGroups.Retain.PublicMembers.Any())
                {
                    foreach (var publicMember in declarationGroups.Retain.PublicMembers)
                    {
                        var destinationRefs = publicMember.References.Where(rf => declarationGroups.Moving.Contains(rf.ParentScoping));
                        foreach (var rf in destinationRefs)
                        {
                            tempRewriter.Replace(rf.Context, $"{_scenario.MoveDefinition.Source.ModuleName}.{rf.IdentifierName}");
                        }
                    }
                }

                return tempRewriter;
            }

            var identifierRefs = ReferencesToPassAsArguments(member).OrderBy(id => id.IdentifierName);

            foreach (var identifierRef in identifierRefs)
            {
                var newArg = $"{MoveMemberResources.Prefix_Parameter}{identifierRef.IdentifierName}";
                if (identifierRef.Declaration.DeclarationType.HasFlag(DeclarationType.Function))
                {
                    if (identifierRef.Context.TryGetAncestor(out VBAParser.IndexExprContext idxExpr))
                    {
                        tempRewriter.Replace(idxExpr, newArg);
                    }
                    else
                    {
                        tempRewriter.Replace(identifierRef.Context, newArg);
                    }
                }
                else if ((identifierRef.Declaration.AsTypeDeclaration?.DeclarationType.HasFlag(DeclarationType.Module) ?? false)
                    && (identifierRef.Declaration.AsTypeDeclaration?.QualifiedModuleName.Equals(_scenario.DestinationContentProvider.QualifiedModuleName) ?? false)
                    && identifierRef.Context.TryGetAncestor<VBAParser.MemberAccessExprContext>(out _))
                {
                    tempRewriter.RemoveMemberAccess(identifierRef);
                }
                else
                {
                    tempRewriter.Replace(identifierRef.Context, newArg);
                }
            }

            tempRewriter.ReplaceDescendentContext<VBAParser.ArgListContext>(member, $"({DestinationSignatureParameters(member)})");
            return tempRewriter;
        }

        private string DestinationNonMemberCodeBlockDefault(Declaration nonMember)
        {
            var newContentRewriter = _rewritingManager.CheckOutTemporaryRewriter(_scenario.SourceContentProvider.QualifiedModuleName);
            if (nonMember.IsVariable() && !_moveDeclarationGroups.IsPropertyBackingVariable(nonMember))
            {
                var variableStmt = nonMember.Context.GetAncestor<VBAParser.VariableStmtContext>();
                Debug.Assert(variableStmt != null);

                var visibility = variableStmt.GetChild<VBAParser.VisibilityContext>().GetText();
                return $"{visibility} {newContentRewriter.GetText(nonMember.Context.Start.TokenIndex, nonMember.Context.Stop.TokenIndex)}";
            }

            if (_moveDeclarationGroups.IsPropertyBackingVariable(nonMember))
            {
                return $"{Tokens.Private} {newContentRewriter.GetText(nonMember.Context.Start.TokenIndex, nonMember.Context.Stop.TokenIndex)}";
            }

            if (nonMember.IsConstant())
            {
                var constStmt = nonMember.Context.GetAncestor<VBAParser.ConstStmtContext>();
                Debug.Assert(constStmt != null);

                var visibility = _scenario.IsOnlyReferencedByMovedElements(nonMember) ? Tokens.Private : Tokens.Public;
                return $"{visibility} {Tokens.Const} {newContentRewriter.GetText(nonMember.Context.Start.TokenIndex, nonMember.Context.Stop.TokenIndex)}";
            }
            return string.Empty;
        }

        private IEnumerable<IdentifierReference> ReferencesToPassAsArguments(Declaration member)
            => _moveMemberContentInfo.ReferenceCandidiatesToPassAsArguments(member);

        private string DestinationSignatureParameters(Declaration member)
            => _moveMemberContentInfo.DestinationSignatureParameters(member);
    }
}

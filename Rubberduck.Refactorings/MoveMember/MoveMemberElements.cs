using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public enum MoveGroups
    {
        Moving,
        SelectedElements,
        Participants,
        SupportingElements,
        AnalyzeableElements,
        Retain,
        MoveAndDelete,
        Forward,
        Remove,
    }

    public interface IProvideMoveDeclarationGroups
    {
        MoveElementGroups Moving { get; }
        MoveElementGroups SelectedElements { get; }
        IEnumerable<Declaration> MoveableElements { set; get; }
        MoveElementGroups Participants { get; }
        MoveElementGroups SupportingElements { get; }
        MoveElementGroups Retain { get; }
        Dictionary<Declaration, string> VariableReferenceReplacement { set; get; }
        MoveElementGroups MoveAndDelete { get; }
        IEnumerable<Declaration> Forward { get; }
        IEnumerable<Declaration> Remove { get; }
        IEnumerable<Declaration> AllDeclarations { get; }
        bool ForwardSelectedMemberCalls { set; get; }
        bool IsPropertyBackingVariable(Declaration variable);
        bool TryGetPropertiesFromBackingVariable(Declaration variable, out List<Declaration> properties);
        bool InternalStateUsedByMovedMembers { get; }
        int CountOfModuleInstanceVariables(Declaration module);
        bool IsSingleDeclarationSelection { get; }
    }

    public interface IMoveMemberRefactoringGroups
    {
        MoveElementGroups MoveAndDelete { get; }
        MoveElementGroups Retain { get; }
        IEnumerable<Declaration> Remove { get; }
        IEnumerable<Declaration> Forward { get; }
    }

    public class MoveMemberDeclarationGroups : IProvideMoveDeclarationGroups
    {
        private MoveMemberPropertiesData PropertiesData { set; get; }
        public MoveElementGroups SelectedElements { get; }
        public IEnumerable<Declaration> AllDeclarations { set; get; }
        private Declaration _module;

        public MoveMemberDeclarationGroups(Declaration module, IEnumerable<Declaration> selectedElements, IDeclarationFinderProvider declarationFinderProvider)
            : this(module, selectedElements, declarationFinderProvider?.DeclarationFinder.Members(module).Where(d => !d.Equals(module)) ?? Enumerable.Empty<Declaration>())
        {
        }

        public MoveMemberDeclarationGroups(MoveDefinition moveDefinition, IDeclarationFinderProvider declarationFinderProvider)
            : this(moveDefinition.Source.Module, moveDefinition.SelectedElements, declarationFinderProvider?.DeclarationFinder.Members(moveDefinition.Source.Module).Where(d => !d.Equals(moveDefinition.Source.Module)) ?? Enumerable.Empty<Declaration>())
        {
            if (moveDefinition.Endpoints.Equals(MoveEndpoints.StdToStd) && !InternalStateUsedByMovedMembers)
            {
                ForwardSelectedMemberCalls = false;
            }
        }

        public MoveMemberDeclarationGroups(Declaration module, IEnumerable<Declaration> selectedElements, IEnumerable<Declaration> moduleDeclarations)
        {
            _module = module;

            SelectedElements = new MoveElementGroups(selectedElements);

            AllDeclarations = moduleDeclarations;

            MoveableElements = moduleDeclarations.Where(element => IsMoveableDeclaration(element));

            PropertiesData = new MoveMemberPropertiesData(MoveableElements, module);

            var allParticipants = DetermineParticipatingDeclarations();

            Participants = new MoveElementGroups(allParticipants);

            SupportingElements = new MoveElementGroups(allParticipants.Except(SelectedElements.AllDeclarations));

            var exclusiveConstants = SupportingElements.NonMembers.Where(nm => nm.IsConstant())
                    .Where((se => se.References.All(rf => selectedElements.Where(sm => sm.IsMember()).Contains(rf.ParentScoping)) //.Equals(theSelectedFunction))
                        || se.References.All(rf => SupportingElements.Members.Contains(rf.ParentScoping))));

            var nonExclusiveConstants = exclusiveConstants.Except(Participants.NonMembers.Where(nm => nm.IsConstant()));

            //From SingleFunctionToStdModules
            var exclusiveVariables = SupportingElements.NonMembers
                    .Where((se => se.References.All(rf => selectedElements.Where(sm => sm.IsMember()).Contains(rf.ParentScoping)) //.Equals(theSelectedFunction))
                        || se.References.All(rf => SupportingElements.Members.Contains(rf.ParentScoping))));

            var nonExclusiveVariables = exclusiveVariables.Except(Participants.NonMembers);

            var allMembers = SupportingElements.Members.Concat(SelectedElements.Members);

            var exclusiveMembers = SupportingElements.Members
                .Where(se => se.References.All(seRefs => allMembers.Contains(seRefs.ParentScoping)));

            var nonExclusiveMembers = exclusiveMembers.Except(Participants.Members);

            var unmoveableDeclarations = nonExclusiveMembers.Except(exclusiveMembers)
                        .Concat(nonExclusiveVariables.Except(exclusiveVariables));

            //Exclusive MemberVariables and Constants
            var privateNonMembersExclusiveToSelectedOrPrivateMovedElements =
                Participants.PrivateNonMembers
                    .Where(p => p.References
                        .All(rf => SelectedElements.Contains(rf.ParentScoping)
                                || Participants.PrivateMembers.Contains(rf.ParentScoping)));

            //No Equivalent
            var privateTypesExclusiveToSelectedOrPrivateMovedElements =
                Participants.PrivateTypeDefinitions
                    .Where(p => p.References
                        .All(rf => SelectedElements.Contains(rf.ParentScoping)
                                || Participants.PrivateMembers.Contains(rf.ParentScoping)));

            //exclusiveMembers
            var privateMembersExclusiveToSelectedOrPrivateMovedElements =
                Participants.PrivateMembers
                    .Where(p => p.References.
                        All(rf => SelectedElements.Contains(rf.ParentScoping) 
                            || Participants.PrivateMembers.Contains(rf.ParentScoping)));

            //Not used
            //var privateNonMembersExclusiveToPublicRetainedElements = 
            //    Participants.PrivateNonMembers
            //        .Where(p => p.References
            //            .All(rf => Retain.PublicMembers.Where(r => Participants.AllDeclarations.Contains(p)).Contains(rf.ParentScoping)));

            var backingVariables = 
                Participants.PrivateNonMembers
                    .Where(p => PropertiesData.IsPropertyBackingVariable(p));

            var backingVariablesToMove = new List<Declaration>();
            foreach (var backingVariable in backingVariables)
            {
                if (PropertiesData.TryGetPropertiesFromBackingVariable(backingVariable, out var properties))
                {
                    if (properties.Any(p => SelectedElements.Contains(p)))
                    {
                        backingVariablesToMove.Add(backingVariable);
                    }
                }
            }

            var elementsToMove = SelectedElements.AllDeclarations
                .Concat(privateTypesExclusiveToSelectedOrPrivateMovedElements)
                //.Concat(privateNonMembersExclusiveToSelectedOrPrivateMovedElements)
                .Concat(exclusiveVariables)
                .Concat(exclusiveConstants)
                //.Concat(privateMembersExclusiveToSelectedOrPrivateMovedElements)
                .Concat(exclusiveMembers)
                .Concat(backingVariablesToMove)
                .Distinct();

            foreach (var nonMember in backingVariablesToMove)
            {
                if (PropertiesData.TryGetPropertiesFromBackingVariable(nonMember, out var properties))
                {
                    VariableReferenceReplacement.Add(nonMember, properties.First().IdentifierName);
                }
            }

            Retain = new MoveElementGroups(moduleDeclarations.Except(elementsToMove).Except(VariableReferenceReplacement.Keys));

            Moving = new MoveElementGroups(elementsToMove);
        }

        public MoveElementGroups Moving { get; }

        public bool IsPropertyBackingVariable(Declaration variable) => PropertiesData.IsPropertyBackingVariable(variable);

        public bool TryGetPropertiesFromBackingVariable(Declaration variable, out List<Declaration> properties)
        {
            return PropertiesData.TryGetPropertiesFromBackingVariable(variable, out properties);
        }

        public bool IsSingleDeclarationSelection
        {
            get
            {
                if (SelectedElements.AllDeclarations.Count() == 1)
                {
                    return true;
                }
                if (SelectedElements.Members.Count() == 2 
                    && SelectedElements.Members.All(m => m.DeclarationType.HasFlag(DeclarationType.Property)))
                {
                    return true;
                }
                return false;
            }
        }

        private bool? _internalStateUsedByMovedMembers;
        public bool InternalStateUsedByMovedMembers
        {
            get
            {
                if (!_internalStateUsedByMovedMembers.HasValue)
                {
                    var allInstanceNonMembers = AllDeclarations
                        .Where(el => !el.IsMember()
                            && !(el.IsLocalVariable() || el.IsLocalConstant())
                            && !el.DeclarationType.HasFlag(DeclarationType.Parameter));

                    var stateContent = new List<Declaration>();
                    foreach (var nonMember in allInstanceNonMembers)
                    {
                        if (TryGetPropertiesFromBackingVariable(nonMember, out var properties))
                        {
                            if (!SelectedElements.Members.All(m => properties.Contains(m)))
                            {
                                stateContent.Add(nonMember);
                            }
                        }
                        else
                        {
                            stateContent.Add(nonMember);
                        }
                    }
                    _internalStateUsedByMovedMembers = stateContent.AllReferences().Any(rf => Participants.Members.Contains(rf.ParentScoping));
                }
                return _internalStateUsedByMovedMembers.Value;
            }
        }

        public IEnumerable<Declaration> MoveableElements { set; get; } = new List<Declaration>();

        public MoveElementGroups Participants { get; }

        public MoveElementGroups SupportingElements { get; }

        public MoveElementGroups Retain { get; }

        public Dictionary<Declaration, string> VariableReferenceReplacement { set; get; } = new Dictionary<Declaration, string>();

        public MoveElementGroups MoveAndDelete
            => new MoveElementGroups(ForwardSelectedMemberCalls ? Moving.Except(Forward) : Moving.AllDeclarations);

        public IEnumerable<Declaration> Forward
            => ForwardSelectedMemberCalls ? SelectedElements.Members : Enumerable.Empty<Declaration>();

        public bool ForwardSelectedMemberCalls { set; get; } = false; // true;

        public IEnumerable<Declaration> Remove => MoveAndDelete.Concat(VariableReferenceReplacement.Keys);

        public int CountOfModuleInstanceVariables(Declaration module)
        {
            if (module is null) { return 0; }

            return AllDeclarations.Where(el => el.IsVariable()
                && (el.AsTypeDeclaration?.IdentifierName.Equals(module.IdentifierName) ?? false)).Count();
        }

        private IEnumerable<Declaration> DetermineParticipatingDeclarations()
        {
            var participatingDeclarations = new HashSet<Declaration>();

            foreach (var element in SelectedElements.AllDeclarations)
            {
                participatingDeclarations.Add(element);
            }

            var allReferences = AllDeclarations.AllReferences().ToList();

            var guard = 0;
            var newElements = SelectedElements.AllDeclarations;
            while (newElements.Any() && guard++ < 100)
            {
                newElements = allReferences
                        .Where(rf => newElements.Contains(rf.ParentScoping)
                                && rf.ParentScoping != rf.Declaration)
                        .Select(rf => rf.Declaration)
                .ToList();

                foreach (var element in newElements)
                {
                    if (IsMoveableDeclaration(element))
                    {
                        participatingDeclarations.Add(element);
                    }
                }
            }
            Debug.Assert(guard < 100);
            if (guard >= 100)
            {
                throw new MoveMemberUnsupportedMoveException(SelectedElements.FirstOrDefault());
            }

            var supportAsTypes = participatingDeclarations.Where(apd => apd.AsTypeDeclaration != null)
                .Select(atd => atd.AsTypeDeclaration.DeclarationType);

            if (supportAsTypes.Any())
            {
                var supports = AllDeclarations.Where(m => supportAsTypes.Contains(m.DeclarationType));
                foreach (var support in supports)
                {
                    participatingDeclarations.Add(support);
                }
            }
            participatingDeclarations.Remove(_module);

            return participatingDeclarations;
        }

        private static bool IsMoveableDeclaration(Declaration declaration)
        {
            return (declaration.DeclarationType.HasFlag(DeclarationType.Member)
                        || declaration.DeclarationType.HasFlag(DeclarationType.Variable) && !declaration.IsLocalVariable()
                        || declaration.DeclarationType.HasFlag(DeclarationType.Constant) && !declaration.IsLocalConstant());
                    //&& !declaration.IsLocalConstant()
                    //&& !declaration.IsLocalVariable();
        }
    }

}

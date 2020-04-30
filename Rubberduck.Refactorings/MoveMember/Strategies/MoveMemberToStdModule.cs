using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Rename;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public interface IMoveMemberRefactoringStrategy
    {
        void RefactorRewrite(MoveMemberModel model, IRewriteSession rewriteSession, IRewritingManager rewritingManager, INewContentProvider contentProvider, out string newModuleContent);
        bool IsApplicable(MoveMemberModel model);
        bool IsExecutableModel(MoveMemberModel model, out string nonExecutableMessage);
    }

    public class MoveMemberToStdModule : MoveMemberStrategyBase, IMoveMemberRefactoringStrategy
    {
        public MoveMemberToStdModule(IDeclarationFinderProvider declarationFinderProvider,
                                        RenameCodeDefinedIdentifierRefactoringAction renameAction,
                                        IMoveMemberMoveGroupsProviderFactory moveGroupsProviderFactory,
                                        IConflictDetectionSessionFactory namingToolsSessionFactory,
                                        IConflictDetectionDeclarationProxyFactory declarationProxyFactory)
            : base(declarationFinderProvider,
                                        renameAction,
                                        moveGroupsProviderFactory,
                                        namingToolsSessionFactory,
                                        declarationProxyFactory)
        { }

        public override bool IsApplicable(MoveMemberModel model)
        {
            //Note: A StandardModule is the default
            //model.Destination.IsStandardModule returns true for non-specified destinations
            if (!model.Destination.IsStandardModule) { return false; }

            var moveGroups = _moveGroupsProviderFactory.Create(model.MoveableMemberSets);

            //Strategy does not support Private Fields as a SelectedDeclaration.  But, the strategy
            //is still applicable if ALL of the Selected Fields are contained in the MoveGroup.Support_Exclusive.
            //Private/MoveGroup.Support_Exclusive are always moved along with the moved Members that reference/depend on them
            if (moveGroups.SelectedPrivateFields.Except(moveGroups.Declarations(MoveGroup.Support_Exclusive)).Any())
            {
                return false;
            }

            if (model.Source.IsStandardModule)
            {
                return TrySetDispositionGroupsForStandardModuleSource(model, moveGroups, out _);
            }

            return TrySetDispositionGroupsForClassModuleSource(model, moveGroups, out _);
        }

        public bool IsExecutableModel(MoveMemberModel model, out string nonExecutableMessage)
        {
            return base.IsExecutableModelBase(model, out nonExecutableMessage);
        }

        protected override Dictionary<MoveDisposition, List<Declaration>> DetermineDispositionGroups(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups)
        {
            var dispositions = EmptyDispositions();
            if (model.Source.IsStandardModule
                && TrySetDispositionGroupsForStandardModuleSource(model, moveGroups, out dispositions))
            {
                return dispositions;
            }

            if (TrySetDispositionGroupsForClassModuleSource(model, moveGroups, out dispositions))
            {
                return dispositions;
            }

            throw new MoveMemberUnsupportedMoveException();
        }

        protected override INewContentProvider LoadSourceNewContentProvider(INewContentProvider contentProvider, MoveMemberModel model) => contentProvider.ResetContent();

        private bool TrySetDispositionGroupsForClassModuleSource(MoveMemberModel model, IMoveMemberGroupsProvider moveGroups, out Dictionary<MoveDisposition, List<Declaration>> dispositions)
        {
            dispositions = EmptyDispositions();

            var participatingMembers = moveGroups.MoveableMemberSets(MoveGroup.AllParticipants)
                        .Where(d => d.Member.IsMember())
                        .SelectMany(pm => pm.Members);

            if (ParticipantsRaiseAnEvent(participatingMembers)
                || ParticipantsIncludeAnEventHandler(participatingMembers, model)
                || ParticipantsIncludeAnInterfaceImplementingMember(participatingMembers, model))
            {
                return false;
            }

            if (moveGroups.Declarations(MoveGroup.Support_NonExclusive).Any()) { return false; }

            //Strategy does not move Public Fields of ObjectTypes.
            //As a Public Field of a StandardModule, we lose control of instantiation logic
            if (model.SelectedDeclarations.Any(nm => nm.DeclarationType.HasFlag(DeclarationType.Variable)
                                    && nm.IsObject)) { return false; }


            //If there are external references to the selected element(s) and the source is a 
            //ClassModule or  UserForm, then they are not be supported by this strategy.
            if (model.SelectedDeclarations.AllReferences().Any(rf => rf.QualifiedModuleName != model.Source.QualifiedModuleName))
            {
                return false;
            }

            dispositions[MoveDisposition.Move] = (moveGroups.Declarations(MoveGroup.Selected)
                .Concat(moveGroups.Declarations(MoveGroup.Support_Exclusive))).ToList();

            dispositions[MoveDisposition.Retain] = (moveGroups.Declarations(MoveGroup.AllParticipants)
                                                        .Except(dispositions[MoveDisposition.Move])).ToList();
            return true;
        }

        private bool ParticipantsIncludeAnInterfaceImplementingMember(IEnumerable<Declaration> participatingMembers, MoveMemberModel model)
        {
            var interfaceImplementingMembers = _declarationFinderProvider.DeclarationFinder.FindAllInterfaceImplementingMembers()
                .Where(ifm => ifm.QualifiedModuleName == model.Source.QualifiedModuleName);

            return participatingMembers.Intersect(interfaceImplementingMembers).Any();
        }

        private static bool ParticipantsRaiseAnEvent(IEnumerable<Declaration> participatingMembers)
                    => participatingMembers.Where(m => m.Context.GetDescendent<VBAParser.RaiseEventStmtContext>() != null).Any();

        private bool ParticipantsIncludeAnEventHandler(IEnumerable<Declaration> participatingMembers, MoveMemberModel model)
        {
            var eventHandlers = _declarationFinderProvider.DeclarationFinder.FindEventHandlers()
                .Where(evh => evh.QualifiedModuleName == model.Source.QualifiedModuleName);

            return participatingMembers.Intersect(eventHandlers).Any();
        }
    }
}

using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;


namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberModel : IRefactoringModel
    {
        public IDeclarationFinderProvider DeclarationFinderProvider => State;
        public RubberduckParserState State { get; }

        public MoveMemberModel(RubberduckParserState state, IRewritingManager rewritingManager)
        {
            State = state;
            MoveRewritingManager = new MoveMemberRewritingManager(rewritingManager);
        }

        public void DefineMove(Declaration member, Declaration destinationModule = null)
        {
            DefiningMember = member;
            CurrentScenario = CreateMoveScenario(DeclarationFinderProvider, member, new MoveDefinitionEndpoint(destinationModule));
            SetStrategy();
        }

        public void DefineMove(Declaration member, string destinationModuleName, ComponentType destinationType)
        {
            DefiningMember = member;
            CurrentScenario = CreateMoveScenario(DeclarationFinderProvider, member, new MoveDefinitionEndpoint(destinationModuleName, destinationType));
            SetStrategy();
        }

        private void SetStrategy()
        {
            Strategy = null;

            var groups = CurrentScenario as IProvideMoveDeclarationGroups;

            //if (MoveMemberStrategy_Common.IsUnsupportedMove_General(CurrentScenario, groups)) { return; }

            //if (MoveMemberStrategy_Common.IsUnsupportedMove_General_Member(CurrentScenario, groups)) { return; }

            var strategies = MoveMemberStrategyProvider.FindStrategies(CurrentScenario, MoveRewritingManager);
            if (strategies.Count() == 1)
            {
                Strategy = strategies.Single();
                return;
            }
            //TODO: display info for 0 and >1 cases
        }

        public IMoveMemberRefactoringStrategy Strategy { private set; get; }

        public void ChangeDestination(string destinationModuleName)
        {
            var destinations = DeclarationFinderProvider.DeclarationFinder.MatchName(destinationModuleName)
                .Where(d => d.DeclarationType.HasFlag(DeclarationType.Module) && d.IsUserDefined);

            DefineMove(CurrentScenario.SelectedElements.First(), destinations.Single());
        }

        private IDictionary<QualifiedModuleName, Declaration> _defaultMembers;
        private IDictionary<QualifiedModuleName, Declaration> DefaultMembers
        {
            get
            {
                if (_defaultMembers is null)
                {
                    _defaultMembers = new Dictionary<QualifiedModuleName, Declaration>();
                    var modules = DeclarationFinderProvider.DeclarationFinder.AllDeclarations.Where(d => d.DeclarationType.HasFlag(DeclarationType.Module) && d.IsUserDefined);

                    foreach (var module in modules)
                    {
                        _defaultMembers.Add(module.QualifiedModuleName, MoveMemberDefaultMember(module)); //, DeclarationFinderProvider));
                    }
                }
                return _defaultMembers;
            }
        }

        private Declaration MoveMemberDefaultMember(Declaration sourceModule) //, IDeclarationFinderProvider _declarationFinderProvider)
        {
            if (sourceModule is null)
            {
                return null;
            }

            var members = DeclarationFinderProvider.DeclarationFinder.Members(sourceModule)
                .Where(m => m.IsMember());

            if (sourceModule.QualifiedModuleName.ComponentType == ComponentType.ClassModule)
            {
                members = members.Except(members.Where(m => MoveMemberResources.IsOrNamedLikeALifeCycleHandler(m)));
            }

            if (members.Count() <= 1)
            {
                return members.FirstOrDefault();
            }

            var publicMembers = members.Where(m => !m.HasPrivateAccessibility()).OrderBy(m => m.IdentifierName);
            if (publicMembers.Any())
            {
                return publicMembers.OrderBy(m => m.IdentifierName).First();
            }
            return members.OrderBy(m => m.IdentifierName).FirstOrDefault();
        }

        public Declaration DefaultMemberToMove(QualifiedModuleName qmn)
            => DefaultMembers.ContainsKey(qmn) ? DefaultMembers[qmn] : null;

        public MoveMemberRewritingManager MoveRewritingManager { get; }

        public IMoveScenario CurrentScenario { set; get; } = MoveScenario.NullMove();

        public string PreviewDestination()
        {
            SetStrategy();
            return Strategy?.PreviewDestination() ?? string.Empty;
        }

        public bool IsValidMoveDefinition => CurrentScenario.IsValidMoveDefinition;

        public Declaration SourceModule => CurrentScenario.SourceContentProvider.Module;

        public Declaration DestinationModule => CurrentScenario.DestinationContentProvider.Module;

        public string DestinationModuleName => CurrentScenario.DestinationContentProvider.ModuleName;

        public bool CreatesNewModule
            => DestinationModule is null && !string.IsNullOrEmpty(DestinationModuleName);

        public Declaration DefiningMember { private set; get; }

        private static Dictionary<MoveDefinition, IMoveScenario> Scenarios { get; } = new Dictionary<MoveDefinition, IMoveScenario>();

        public static IMoveScenario CreateMoveScenario(IDeclarationFinderProvider declarationFinderProvider, Declaration selectedDeclaration, MoveDefinitionEndpoint destinationEndpoint)
        {
            if (selectedDeclaration is null)
            {
                return MoveScenario.NullMove();
            }

            var sourceModule = declarationFinderProvider.DeclarationFinder.ModuleDeclaration(selectedDeclaration.QualifiedModuleName);

            var selectedDeclarations = new List<Declaration>() { selectedDeclaration };

            if (selectedDeclaration.DeclarationType.HasFlag(DeclarationType.Property))
            {
                selectedDeclarations = (from dec in declarationFinderProvider.DeclarationFinder.Members(sourceModule)
                                 where dec.DeclarationType.HasFlag(DeclarationType.Property)
                                         && dec.IdentifierName.Equals(selectedDeclaration.IdentifierName)
                                 orderby dec.Selection
                                 select dec).ToList();
            }

            var moveDefinition = new MoveDefinition(new MoveDefinitionEndpoint(sourceModule), destinationEndpoint, selectedDeclarations);

            if (Scenarios.TryGetValue(moveDefinition, out IMoveScenario scenario))
            {
                return scenario;
            }

            scenario = new MoveScenario(moveDefinition, declarationFinderProvider);

            Scenarios.Add(moveDefinition, scenario);
            return scenario;
        }
    }
}

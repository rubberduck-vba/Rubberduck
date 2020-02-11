using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;


namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberModel : IRefactoringModel
    {
        private readonly Func<MoveMemberModel, string> _previewDelegate;
        public IDeclarationFinderProvider DeclarationFinderProvider { get; }
        private List<Declaration> _selectedDeclarations;

        public MoveMemberModel(IDeclarationFinderProvider declarationFinderProvider, IRewritingManager rewritingManager, Func<MoveMemberModel, string> previewDelegate)
        {
            MoveRewritingManager = rewritingManager;
            _previewDelegate = previewDelegate;
            DeclarationFinderProvider = declarationFinderProvider;
            _selectedDeclarations = new List<Declaration>();
        }

        public void DefineMove(Declaration member, Declaration destinationModule = null)
        {
            DefiningMember = member;
            InitializeSelectedDeclarations(member);
            var sourceModule = DeclarationFinderProvider.DeclarationFinder.ModuleDeclaration(member.QualifiedModuleName);
            Source = new MoveSource(new MoveMemberEndpoint(sourceModule));
            Destination = new MoveDestination(new MoveMemberEndpoint(destinationModule));

            MoveGroups = new MoveMemberGroups(SelectedDeclarations, DeclarationFinderProvider);
        }

        public void DefineMove(Declaration member, string destinationModuleName, ComponentType destinationType = ComponentType.StandardModule)
        {
            DefiningMember = member;
            InitializeSelectedDeclarations(member);
            var sourceModule = DeclarationFinderProvider.DeclarationFinder.ModuleDeclaration(member.QualifiedModuleName);
            Source = new MoveSource(new MoveMemberEndpoint(sourceModule));
            Destination = new MoveDestination(new MoveMemberEndpoint(destinationModuleName, destinationType));

            MoveGroups = new MoveMemberGroups(SelectedDeclarations, DeclarationFinderProvider);
        }

        public IEnumerable<Declaration> SelectedDeclarations => _selectedDeclarations;

        public void AddDeclarationToMove(Declaration declaration)
        {
            _selectedDeclarations.Add(declaration);
            MoveGroups = new MoveMemberGroups(_selectedDeclarations, DeclarationFinderProvider);
        }

        public void RemoveDeclarationToMove(Declaration declaration)
        {
            _selectedDeclarations.Remove(declaration);
            MoveGroups = new MoveMemberGroups(_selectedDeclarations, DeclarationFinderProvider);
        }

        public IMoveMemberGroups MoveGroups { private set; get; }

        public IEnumerable<Declaration> AllSourceModuleDeclarations
            => DeclarationFinderProvider.DeclarationFinder.Members(Source.QualifiedModuleName);

        public IEnumerable<Declaration> AllDestinationModuleDeclarations
            => Destination.IsExistingModule(out var module)
                ? DeclarationFinderProvider.DeclarationFinder.Members(module.QualifiedModuleName)
                : Enumerable.Empty<Declaration>();

        public bool IsStdModuleSource => Source.ComponentType.Equals(ComponentType.StandardModule);

        public bool IsStdModuleDestination => Destination?.ComponentType.Equals(ComponentType.StandardModule) ?? false;

        public IMoveMemberRefactoringStrategy Strategy
        {
            get
            {
                var strategy = new MoveMemberToStdModule();
                if (strategy.IsApplicable(this))
                {
                    return strategy;
                }
                return null;
            }
        }

        public void ChangeDestination(string destinationModuleName)
        {
            var destination = DeclarationFinderProvider.DeclarationFinder.MatchName(destinationModuleName)
                .Where(d => d.DeclarationType.HasFlag(DeclarationType.Module) && d.IsUserDefined).SingleOrDefault();

            if (destination is null)
            {
                DefineMove(SelectedDeclarations.First(), destinationModuleName);
                return;
            }
            DefineMove(SelectedDeclarations.First(), destination);
        }

        public IRewritingManager MoveRewritingManager { get; }

        public string PreviewDestination()
        {
            if (_previewDelegate is null)
            {
                return string.Empty;
            }

            return _previewDelegate(this);
        }

        public bool HasValidDestination
        {
            get
            {
                return !(Destination.ModuleName.Equals(Source.ModuleName)
                    || IsInvalidDestinationModuleName());
            }
        }

        private bool IsInvalidDestinationModuleName()
        {
            if (string.IsNullOrEmpty(Destination.ModuleName))
            {
                return false;
            }

            return VBAIdentifierValidator.TryMatchInvalidIdentifierCriteria(Destination.ModuleName, DeclarationType.Module, out var criteriaMatchMessage);
        }

        public string DestinationModuleName => Destination.ModuleName;

        public bool CreatesNewModule
            => !Destination.IsExistingModule(out _) && !string.IsNullOrEmpty(Destination.ModuleName);

        public Declaration DefiningMember { private set; get; }

        public IMoveSource Source { private set; get; }

        public IMoveDestination Destination { private set; get; }

        private void InitializeSelectedDeclarations(Declaration member)
        {
            _selectedDeclarations = new List<Declaration>() { member };
            if (member.DeclarationType.HasFlag(DeclarationType.Property))
            {
                var sourceModule = DeclarationFinderProvider.DeclarationFinder.ModuleDeclaration(member.QualifiedModuleName);
                _selectedDeclarations = (from dec in DeclarationFinderProvider.DeclarationFinder.Members(sourceModule)
                                         where dec.DeclarationType.HasFlag(DeclarationType.Property)
                                                 && dec.IdentifierName.Equals(member.IdentifierName)
                                         orderby dec.Selection
                                         select dec).ToList();
            }
            return;
        }
    }
}

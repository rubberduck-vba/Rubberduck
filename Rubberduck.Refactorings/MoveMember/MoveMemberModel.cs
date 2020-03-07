using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.Resources;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;


namespace Rubberduck.Refactorings.MoveMember
{
    public enum PreviewModule { Source, Destination}

    public class MoveMemberModel : IRefactoringModel
    {
        private readonly IMoveMemberObjectsFactory _moveMemberFactory;
        private Dictionary<string, IMoveableMemberSet> _moveablesByName;
        public IDeclarationFinderProvider DeclarationFinderProvider { get; }

        public MoveMemberModel(Declaration target, IDeclarationFinderProvider declarationFinderProvider)
        {
            DeclarationFinderProvider = declarationFinderProvider;

            _moveMemberFactory = new MoveMemberObjectsFactory(declarationFinderProvider);

            Source = _moveMemberFactory.CreateMoveSourceProxy(target);

            Destination = _moveMemberFactory.CreateMoveDestinationProxy(null);

            _moveablesByName = _moveMemberFactory.CreateMoveables(target).ToDictionary(mm => mm.IdentifierName);

            _previewDelegate = null;
        }

        public IMoveSourceModuleProxy Source { private set; get; }

        public IMoveDestinationModuleProxy Destination { private set; get; }

        public IReadOnlyCollection<IMoveableMemberSet> MoveableMembers => _moveablesByName.Values;

        public IMoveableMemberSet MoveableMemberSetByName(string identifier) => _moveablesByName[identifier];

        public IEnumerable<Declaration> SelectedDeclarations => MoveableMembers
                                            .Where(mc => mc.IsSelected)
                                            .SelectMany(selected => selected.Members);

        public IMoveMemberObjectsFactory MoveMemberFactory => _moveMemberFactory;

        public void ChangeDestination(string destinationModuleName, ComponentType componentType = ComponentType.StandardModule)
        {
            Destination = _moveMemberFactory.CreateMoveDestination(destinationModuleName, componentType);
        }

        public void ChangeDestination(Declaration destinationModule)
        {
            Destination = _moveMemberFactory.CreateMoveDestinationProxy(destinationModule);
        }

        public string PreviewModuleContent(PreviewModule previewModule)
        {
            if (_previewDelegate is null)
            {
                return string.Empty;
            }

            return _previewDelegate(this, previewModule);
        }

        private Func<MoveMemberModel, PreviewModule, string> _previewDelegate;
        public Func<MoveMemberModel, PreviewModule, string> PreviewBuilder
        {
            set => _previewDelegate = value;
            get => _previewDelegate;
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
    }
}

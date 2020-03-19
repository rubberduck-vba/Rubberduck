using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Support = RubberduckTests.Refactoring.MoveMember.MoveMemberTestSupport;

namespace RubberduckTests.Refactoring.MoveMember
{
    public struct ModuleDefinition
    {
        public ModuleDefinition(string name, ComponentType compType, string content = null)
        {
            ModuleName = name;
            ComponentType = compType;
            ModuleContent = content ?? $"{Tokens.Option} {Tokens.Explicit}";
        }

        public string ModuleName { get; }
        public ComponentType ComponentType { get; }
        public string ModuleContent { get; }
        public (string Name, string Content, ComponentType ComponentType) AsTuple
            => (ModuleName, ModuleContent, ComponentType);
    }

    public class TestMoveDefinition
    {
        private Dictionary<string, ModuleDefinition> _moduleDefs;

        public TestMoveDefinition(MoveEndpoints endpoints, (string identifier, DeclarationType declarationType) selection, string sourceContent = null) //, bool createNewModule = false)
        {
            _moduleDefs = new Dictionary<string, ModuleDefinition>();
            _otherSelectedElements = new List<(string, DeclarationType)>();
            CreateNewModule = false;
            Endpoints = endpoints;
            SelectedElement = selection.identifier;
            SelectedDeclarationType = selection.declarationType;
            StrategyName = null;

            if (sourceContent != null)
            {
                SetEndpointContent(sourceContent, null);
            }
        }

        public MoveMemberModel ModelBuilder(IDeclarationFinderProvider declarationFinderProvider)
        {
            var sourceModuleName = SourceModuleName;
            var selectedElement = SelectedElement;
            var target = declarationFinderProvider.DeclarationFinder.DeclarationsWithType(SelectedDeclarationType)
                .Where(t => t.ParentDeclaration.IdentifierName == sourceModuleName)
                .Single(declaration => declaration.IdentifierName == selectedElement);

            var model = new MoveMemberModel(target, declarationFinderProvider);
            foreach ((string ID, DeclarationType decType) in _otherSelectedElements)
            {
                var moveableMemberSet = model.MoveableMemberSetByName(ID);
                moveableMemberSet.IsSelected = true;
            }

            Declaration destination = null;
            if (!CreateNewModule)
            {
                destination = model.DeclarationFinderProvider.DeclarationFinder.DeclarationsWithType(DeclarationType.Module)
                    .Single(t => t.IdentifierName == DestinationModuleName);
                model.ChangeDestination(destination);
                SetTestStrategyName(model);
                return model;
            }

            model.ChangeDestination(DestinationModuleName, DestinationComponentType);

            SetTestStrategyName(model);
            return model;
        }

        private void SetTestStrategyName(MoveMemberModel model)
        {
            StrategyName = null;
            if (MoveMemberObjectsFactory.TryCreateStrategy(model, out var strategy))
            {
                StrategyName = strategy.GetType().Name;
            }
        }

        public string StrategyName { set; get; }
        public MoveEndpoints Endpoints { get; }
        public string SelectedElement { get; }
        public DeclarationType SelectedDeclarationType { get; }

        private List<(string, DeclarationType)> _otherSelectedElements;
        public void AddSelectedDeclaration(string identifier, DeclarationType declarationType)
        {
            _otherSelectedElements.Add((identifier, declarationType));
        }

        private string _sourceModuleName;
        public string SourceModuleName
        {
            set => _sourceModuleName = value;
            get => _sourceModuleName ?? DefaultSourceModuleNameForEndpoint(Endpoints);
        }

        private string _destinationModuleName;
        public string DestinationModuleName
        {
            set => _destinationModuleName = value;
            get => _destinationModuleName ?? DefaultDestinationModuleNameForEndpoint(Endpoints);
        }
        public bool CreateNewModule { set; get; }

        public bool IsClassDestination => Endpoints == MoveEndpoints.ClassToClass || Endpoints == MoveEndpoints.StdToClass;
        public bool IsStdModuleDestination => Endpoints == MoveEndpoints.ClassToStd || Endpoints == MoveEndpoints.StdToStd;
        public bool IsClassSource => Endpoints == MoveEndpoints.ClassToClass || Endpoints == MoveEndpoints.ClassToStd;
        public bool IsFormSource => Endpoints == MoveEndpoints.FormToClass || Endpoints == MoveEndpoints.FormToStd;
        public bool IsStdModuleSource => Endpoints == MoveEndpoints.StdToClass || Endpoints == MoveEndpoints.StdToStd;

        private string DefaultSourceModuleNameForEndpoint(MoveEndpoints endpoints)
        {
            var defaultSourceModuleName = Support.DEFAULT_SOURCE_MODULE_NAME;
            switch (endpoints)
            {
                case MoveEndpoints.ClassToStd:
                    defaultSourceModuleName = Support.DEFAULT_SOURCE_CLASS_NAME;
                    break;
                case MoveEndpoints.FormToStd:
                    defaultSourceModuleName = Support.DEFAULT_SOURCE_FORM_NAME;
                    break;
            }
            return defaultSourceModuleName;
        }

        private string DefaultDestinationModuleNameForEndpoint(MoveEndpoints endpoints)
        {
            return IsStdModuleDestination
                ? Support.DEFAULT_DESTINATION_MODULE_NAME
                : Support.DEFAULT_DESTINATION_CLASS_NAME;
        }

        public ComponentType DestinationComponentType
        {
            get
            {
                switch (Endpoints)
                {
                    case MoveEndpoints.ClassToClass:
                    case MoveEndpoints.StdToClass:
                    case MoveEndpoints.FormToClass:
                        return ComponentType.ClassModule;
                    default:
                        return ComponentType.StandardModule;
                }
            }
        }

        public ComponentType SourceComponentType
        {
            get
            {
                switch (Endpoints)
                {
                    case MoveEndpoints.ClassToStd:
                    case MoveEndpoints.ClassToClass:
                        return ComponentType.ClassModule;
                    case MoveEndpoints.FormToStd:
                    case MoveEndpoints.FormToClass:
                        return ComponentType.UserForm;
                    default:
                        return ComponentType.StandardModule;
                }
            }
        }

        public void Add(ModuleDefinition moduleDef)
        {
            if (!_moduleDefs.TryGetValue(moduleDef.ModuleName, out _))
            {
                _moduleDefs.Add(moduleDef.ModuleName, moduleDef);
            }
        }

        public void SetEndpointContent(string sourceContent, string destinationContent = null)
        {
            Add(LoadSourceModuleContent(sourceContent));
            if (!CreateNewModule)
            {
                Add(LoadDestinationModuleContent(destinationContent));
            }
        }

        public ModuleDefinition[] ModuleDefinitions => _moduleDefs.Values.ToArray();

        public ModuleDefinition LoadSourceModuleContent(string content = null)
            => new ModuleDefinition(SourceModuleName, SourceComponentType, content ?? $"{Tokens.Option} {Tokens.Explicit}");

        public ModuleDefinition LoadDestinationModuleContent(string content = null)
            => new ModuleDefinition(DestinationModuleName, DestinationComponentType, content ?? $"{Tokens.Option} {Tokens.Explicit}");

        public IEnumerable<(string Name, string Content, ComponentType ComponentType)> ModuleTuples
        {
            get
            {
                var allTuples = new List<(string Name, string Content, ComponentType ComponentType)>();
                foreach (var value in ModuleDefinitions)
                {
                    allTuples.Add(value.AsTuple);
                }
                return allTuples;
            }
        }
    }
}

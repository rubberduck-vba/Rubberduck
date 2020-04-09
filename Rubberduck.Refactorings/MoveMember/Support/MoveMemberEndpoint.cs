using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Refactorings.MoveMember
{
    public enum MoveEndpoints
    {
        Undefined,
        StdToStd,
        ClassToStd,
        ClassToClass,
        StdToClass,
        FormToStd,
        FormToClass
    };

    public interface IMoveMemberEndpoint
    {
        string ModuleName { get; }
        ComponentType ComponentType { get; }
        bool IsStandardModule { get; }
        bool IsClassModule { get; }
        bool IsUserFormModule { get; }
        IEnumerable<Declaration> ModuleDeclarations { get; }
        bool TryGetCodeSectionStartIndex(out int insertionIndex);
        Declaration Module { get; }
    }

    public interface IMoveSourceEndpoint : IMoveMemberEndpoint
    {
        QualifiedModuleName QualifiedModuleName { get; }
        IReadOnlyCollection<IMoveableMemberSet> MoveableMembers { get; }
        IMoveableMemberSet MoveableMemberSetByName(string identifier);
    }

    public interface IMoveDestinationEndpoint : IMoveMemberEndpoint
    {
        bool IsExistingModule(out Declaration module);
    }

    public class MoveSourceEndpoint : IMoveSourceEndpoint
    {
        private IMoveMemberEndpoint _endpoint;
        private Dictionary<string, IMoveableMemberSet> _moveablesByName;

        public MoveSourceEndpoint(Declaration target, IMoveMemberEndpoint endpoint, IMoveableMemberSetsFactory moveableMemberSetFactory)
        {
            _endpoint = endpoint;
            _moveablesByName = moveableMemberSetFactory.Create(endpoint.Module).ToDictionary(mm => mm.IdentifierName);
            _moveablesByName[target.IdentifierName].IsSelected = true;
        }

        public IEnumerable<Declaration> ModuleDeclarations
            => _endpoint.ModuleDeclarations;

        public QualifiedModuleName QualifiedModuleName => Module.QualifiedModuleName;
        public Declaration Module => _endpoint.Module;
        public string ModuleName => _endpoint.Module.IdentifierName;
        public ComponentType ComponentType => _endpoint.ComponentType;

        public bool IsStandardModule => _endpoint.IsStandardModule;
        public bool IsClassModule => _endpoint.IsClassModule;
        public bool IsUserFormModule => _endpoint.IsUserFormModule;

        public bool TryGetCodeSectionStartIndex(out int insertionIndex)
            => _endpoint.TryGetCodeSectionStartIndex(out insertionIndex);

        public IReadOnlyCollection<IMoveableMemberSet> MoveableMembers 
            => _moveablesByName.Values;

        public IMoveableMemberSet MoveableMemberSetByName(string identifier) 
            => _moveablesByName[identifier];
    }

    public class MoveDestinationEndpoint : IMoveDestinationEndpoint
    {
        private IMoveMemberEndpoint _endpoint;

        public MoveDestinationEndpoint(IMoveMemberEndpoint endpoint)
        {
            _endpoint = endpoint;
        }

        public string ModuleName => _endpoint.ModuleName;
        public ComponentType ComponentType => _endpoint.ComponentType;

        public IEnumerable<Declaration> ModuleDeclarations
            => _endpoint.ModuleDeclarations;

        //Destination defaults to StandardModule if the Destination module is unassigned
        public bool IsStandardModule => _endpoint.IsStandardModule || string.IsNullOrEmpty(_endpoint.ModuleName);
        public bool IsClassModule => _endpoint.IsClassModule;
        public bool IsUserFormModule => _endpoint.IsUserFormModule;

        public Declaration Module => _endpoint.Module;
        public bool IsExistingModule(out Declaration module)
        {
            module = _endpoint.Module;
            return module != null;
        }

        public bool TryGetCodeSectionStartIndex(out int insertionIndex)
            => _endpoint.TryGetCodeSectionStartIndex(out insertionIndex);
    }

    public class MoveMemberEndpoint : IMoveMemberEndpoint
    {
        public MoveMemberEndpoint(string moduleName, ComponentType moduleComponentType, IDeclarationFinderProvider declarationFinderProvider)
            : this(null, declarationFinderProvider)
        {
            ModuleName = moduleName;
            ComponentType = moduleComponentType;
        }

        public MoveMemberEndpoint(Declaration module, IDeclarationFinderProvider declarationFinderProvider)
        {
            QualifiedModuleName = module?.QualifiedModuleName;
            Module = module;
            ModuleName = module?.IdentifierName ?? string.Empty;
            ComponentType = module?.QualifiedModuleName.ComponentType ?? ComponentType.Undefined;
            DeclarationFinderProvider = declarationFinderProvider;
        }

        public QualifiedModuleName? QualifiedModuleName { private set; get; }
        public Declaration Module { private set; get; }
        public string ModuleName { private set; get; }
        public ComponentType ComponentType { private set; get; }
        public bool IsStandardModule => ComponentType.Equals(ComponentType.StandardModule);
        public bool IsClassModule => ComponentType.Equals(ComponentType.ClassModule);
        public bool IsUserFormModule => ComponentType.Equals(ComponentType.UserForm);
        public IDeclarationFinderProvider DeclarationFinderProvider { private set; get; }

        public IEnumerable<Declaration> ModuleDeclarations
        {
            get
            {
                if (Module != null)
                {
                    return DeclarationFinderProvider.DeclarationFinder.Members(Module);
                }
                return Enumerable.Empty<Declaration>();
            }
        }

        public bool TryGetCodeSectionStartIndex(out int insertionIndex)
        {
            insertionIndex = -1;
            if (Module != null)
            {
                insertionIndex = DeclarationFinderProvider.DeclarationFinder.Members(Module.QualifiedModuleName)
                        .Where(d => d.IsMember() 
                                        && !(d.DeclarationType.Equals(DeclarationType.LibraryFunction)
                                                || d.DeclarationType.Equals(DeclarationType.LibraryProcedure)))
                        .OrderBy(c => c.Selection)
                        .FirstOrDefault()?.Context.Start.TokenIndex ?? -1;
            }
            return insertionIndex > -1;
        }
    }
}

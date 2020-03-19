using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.MoveMember.Extensions;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

    public interface IMoveMemberModuleProxy
    {
        string ModuleName { get; }
        ComponentType ComponentType { get; }
        bool IsStandardModule { get; }
        bool IsClassModule { get; }
        bool IsUserFormModule { get; }
        IEnumerable<Declaration> ModuleDeclarations { get; }
        bool TryGetCodeSectionStartIndex(out int insertionIndex);
    }

    public interface IMoveSourceModuleProxy : IMoveMemberModuleProxy
    {
        Declaration Module { get; }
        QualifiedModuleName QualifiedModuleName { get; }
    }

    public interface IMoveDestinationModuleProxy : IMoveMemberModuleProxy
    {
        bool IsExistingModule(out Declaration module);
    }

    public class MoveSourceModuleProxy : IMoveSourceModuleProxy
    {
        private MoveMemberEndpoint _endpoint;

        public MoveSourceModuleProxy(MoveMemberEndpoint endpoint)
        {
            _endpoint = endpoint;
        }

        public IEnumerable<Declaration> ModuleDeclarations 
                    => _endpoint.DeclarationFinderProvider.DeclarationFinder.Members(_endpoint.Module);

        public QualifiedModuleName QualifiedModuleName => Module.QualifiedModuleName;
        public Declaration Module => _endpoint.Module;
        public string ModuleName => _endpoint.Module.IdentifierName;
        public ComponentType ComponentType => _endpoint.ComponentType;

        public bool IsStandardModule => _endpoint.IsStandardModule;
        public bool IsClassModule => _endpoint.IsClassModule;
        public bool IsUserFormModule => _endpoint.IsUserFormModule;

        public bool TryGetCodeSectionStartIndex(out int insertionIndex)
            => _endpoint.TryGetCodeSectionStartIndex(out insertionIndex);
    }

    public class MoveDestinationModuleProxy : IMoveDestinationModuleProxy
    {
        private MoveMemberEndpoint _endpoint;

        public MoveDestinationModuleProxy(MoveMemberEndpoint endpoint)
        {
            _endpoint = endpoint;
        }

        public string ModuleName => _endpoint.ModuleName;
        public ComponentType ComponentType => _endpoint.ComponentType;

        public IEnumerable<Declaration> ModuleDeclarations
        {
            get
            {
                if (IsExistingModule(out var module))
                {
                    return _endpoint.DeclarationFinderProvider.DeclarationFinder.Members(module);
                }
                return Enumerable.Empty<Declaration>();
            }
        }

        //Destination defaults to StandardModule if the Destination module is unassigned
        public bool IsStandardModule => _endpoint.IsStandardModule || string.IsNullOrEmpty(_endpoint.ModuleName);
        public bool IsClassModule => _endpoint.IsClassModule;
        public bool IsUserFormModule => _endpoint.IsUserFormModule;

        public bool IsExistingModule(out Declaration module)
        {
            module = _endpoint.Module;
            return module != null;
        }

        public bool TryGetCodeSectionStartIndex(out int insertionIndex)
            => _endpoint.TryGetCodeSectionStartIndex(out insertionIndex);
    }

    public struct MoveMemberEndpoint
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

        public QualifiedModuleName? QualifiedModuleName { get; }
        public Declaration Module { get; }
        public string ModuleName { get; }
        public ComponentType ComponentType { get; }
        public bool IsStandardModule => ComponentType.Equals(ComponentType.StandardModule);
        public bool IsClassModule => ComponentType.Equals(ComponentType.ClassModule);
        public bool IsUserFormModule => ComponentType.Equals(ComponentType.UserForm);
        public IDeclarationFinderProvider DeclarationFinderProvider { get; }

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

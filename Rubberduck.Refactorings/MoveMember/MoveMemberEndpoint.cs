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

    public interface IMoveEndpoint
    {
        string ModuleName { get; }
        ComponentType ComponentType { get; }
    }

    public interface IMoveSource : IMoveEndpoint
    {
        Declaration Module { get; }
        QualifiedModuleName QualifiedModuleName { get; }
    }

    public interface IMoveDestination : IMoveEndpoint
    {
        bool IsExistingModule(out Declaration module);
        bool TryGetCodeSectionStartIndex(IDeclarationFinderProvider declarationFinderProvider, out int insertionIndex);
    }

    public class MoveSource : IMoveSource
    {
        private MoveMemberEndpoint _endpoint;

        public MoveSource(MoveMemberEndpoint endpoint)
        {
            _endpoint = endpoint;
        }

        public QualifiedModuleName QualifiedModuleName => Module.QualifiedModuleName;
        public Declaration Module => _endpoint.Module;
        public string ModuleName => _endpoint.Module.IdentifierName;
        public ComponentType ComponentType => _endpoint.ComponentType;
    }

    public class MoveDestination : IMoveDestination
    {
        private MoveMemberEndpoint _endpoint;

        public MoveDestination(MoveMemberEndpoint endpoint)
        {
            _endpoint = endpoint;
        }

        public string ModuleName => _endpoint.ModuleName;
        public ComponentType ComponentType => _endpoint.ComponentType;
        public bool IsExistingModule(out Declaration module)
        {
            module = null;
            if (_endpoint.Module != null)
            {
                module = _endpoint.Module;
                return true;
            }
            return false;
        }

        public bool TryGetCodeSectionStartIndex(IDeclarationFinderProvider declarationFinderProvider, out int insertionIndex)
        {
            insertionIndex = -1;
            if (IsExistingModule(out var module))
            {
                insertionIndex = declarationFinderProvider.DeclarationFinder.Members(module.QualifiedModuleName)
                        .Where(d => d.IsMember())
                        .OrderBy(c => c.Selection)
                        .FirstOrDefault()?.Context.Start.TokenIndex ?? -1;
            }
            return insertionIndex > -1;
        }
    }

    public struct MoveMemberEndpoint
    {
        public MoveMemberEndpoint(string moduleName, ComponentType moduleComponentType)
            : this(null)
        {
            ModuleName = moduleName;
            ComponentType = moduleComponentType;
        }

        public MoveMemberEndpoint(Declaration module)
        {
            QualifiedModuleName = module?.QualifiedModuleName;
            Module = module;
            ModuleName = module?.IdentifierName ?? string.Empty;
            ComponentType = module?.QualifiedModuleName.ComponentType ?? ComponentType.Undefined;
        }

        public QualifiedModuleName? QualifiedModuleName { get; }
        public Declaration Module { get; }
        public string ModuleName { get; }
        public ComponentType ComponentType { get; }
    }
}

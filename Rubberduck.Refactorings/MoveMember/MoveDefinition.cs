using Rubberduck.Common;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
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

    public struct MoveDefinition
    {
        private static Dictionary<(ComponentType, ComponentType), MoveEndpoints> ComponentsToMoveType
            = new Dictionary<(ComponentType, ComponentType), MoveEndpoints>()
            {
                [(ComponentType.ClassModule, ComponentType.ClassModule)] = MoveEndpoints.ClassToClass,
                [(ComponentType.ClassModule, ComponentType.StandardModule)] = MoveEndpoints.ClassToStd,
                [(ComponentType.StandardModule, ComponentType.ClassModule)] = MoveEndpoints.StdToClass,
                [(ComponentType.StandardModule, ComponentType.StandardModule)] = MoveEndpoints.StdToStd,
                [(ComponentType.UserForm, ComponentType.StandardModule)] = MoveEndpoints.FormToStd,
                [(ComponentType.UserForm, ComponentType.ClassModule)] = MoveEndpoints.FormToClass,
                [(ComponentType.ClassModule, ComponentType.Undefined)] = MoveEndpoints.Undefined,
                [(ComponentType.StandardModule, ComponentType.Undefined)] = MoveEndpoints.Undefined,
                [(ComponentType.UserForm, ComponentType.Undefined)] = MoveEndpoints.Undefined,
            };

        private int _hashCode;

        public MoveDefinition(MoveDefinitionEndpoint source, MoveDefinitionEndpoint destination, IEnumerable<Declaration> selectedElements)
        {
            Debug.Assert(source.Module != null);

            Source = source;
            Destination = destination;
            SelectedElements = selectedElements;
            Endpoints = ComponentsToMoveType[(Source.ComponentType, Destination.ComponentType)];
            
            var stringized = $"{Source.ModuleName}, {Destination.ModuleName}, { string.Join(", ", SelectedElements.OrderBy(d => d.IdentifierName).Select(d => d.IdentifierName))}";
            _hashCode = stringized.GetHashCode();
        }

        public MoveDefinitionEndpoint Source { get; }

        public MoveDefinitionEndpoint Destination { get; }

        public MoveEndpoints Endpoints { get; }

        public IEnumerable<Declaration> SelectedElements { get; }

        public bool IsClassModuleSource => Endpoints == MoveEndpoints.ClassToClass || Endpoints == MoveEndpoints.ClassToStd;

        public bool IsClassModuleDestination => Endpoints == MoveEndpoints.ClassToClass || Endpoints == MoveEndpoints.StdToClass || Endpoints == MoveEndpoints.FormToClass;

        public bool IsStdModuleSource => Endpoints == MoveEndpoints.StdToStd || Endpoints == MoveEndpoints.StdToClass;

        public bool IsStdModuleDestination => Endpoints == MoveEndpoints.StdToStd || Endpoints == MoveEndpoints.ClassToStd || Endpoints == MoveEndpoints.FormToStd;

        public bool IsUserFormSource => Endpoints == MoveEndpoints.FormToClass || Endpoints == MoveEndpoints.FormToStd;

        public override int GetHashCode() => _hashCode;

        public override string ToString()
            => $"Source:{Source.ModuleName} Destination:{Destination.Module?.IdentifierName ?? Tokens.Null} Selected: {string.Join(", ",SelectedElements.OrderBy(d => d.IdentifierName))}";

        public override bool Equals(object obj)
        {
            if (!(obj is MoveDefinition moveDef))
            {
                return false;
            }

            if (moveDef.Source.Module != Source.Module
                || (moveDef.Destination.Module != Destination.Module && moveDef.Destination.ModuleName != Destination.ModuleName))
            {
                return false;
            }

            return SelectedElements.All(se => moveDef.SelectedElements.Contains(se));
        }
    }

    public struct MoveDefinitionEndpoint
    {
        public MoveDefinitionEndpoint(string moduleName, ComponentType moduleComponentType)
            : this(null)
        {
            ModuleName = moduleName;
            ComponentType = moduleComponentType;
        }

        public MoveDefinitionEndpoint(Declaration module)
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

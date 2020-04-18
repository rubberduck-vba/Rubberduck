using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.VBEditor.SafeComWrappers;
using System;

namespace RubberduckTests.Refactoring.MoveMember
{
    public static class MoveMemberTestStringExtensions
    {
        public static bool OccursOnce(this string toFind, string content)
        {
            var firstIdx = content.IndexOf(toFind);
            return (firstIdx == content.LastIndexOf(toFind)) && firstIdx > -1;
        }
    }

    public static class MoveEndpointsTestExtensions
    {
        private const string DEFAULT_SOURCE_MODULE_NAME = "DfltSrcStd";
        private const string DEFAULT_SOURCE_CLASS_NAME = "DfltSrcClass";
        private const string DEFAULT_SOURCE_FORM_NAME = "DfltSrcForm";
        private const string DEFAULT_DESTINATION_MODULE_NAME = "DfltDestStd";
        private const string DEFAULT_DESTINATION_CLASS_NAME = "DfltDestClass";

        public static (string, string, ComponentType)[] ToModulesTuples(this MoveEndpoints endpoints, string sourceContent, string destinationContent)
        {
            var sourceTuple = ToSourceTuple(endpoints, sourceContent);
            if (string.IsNullOrEmpty(destinationContent))
            {
                destinationContent = $"Option Explicit{Environment.NewLine}{Environment.NewLine}";
            }
            var destinationTuple = ToDestinationTuple(endpoints, destinationContent);
            return new(string, string, ComponentType)[] { sourceTuple, destinationTuple };
        }

        public static (string moduleName, string content, ComponentType componentType) ToSourceTuple(this MoveEndpoints endpoints, string content)
        {
            var endpointAttributes = endpoints.ToSourceAttributes();
            return (endpointAttributes.ModuleName, content, endpointAttributes.ComponentType);
        }

        public static (string moduleName, string content, ComponentType componentType) ToDestinationTuple(this MoveEndpoints endpoints, string content)
        {
            var endpointAttributes = endpoints.ToDestinationAttributes();
            return (endpointAttributes.ModuleName, content, endpointAttributes.ComponentType);
        }

        public static string SourceModuleName(this MoveEndpoints endpoints)
        {
            return endpoints.ToSourceAttributes().ModuleName;
        }

        public static string DestinationClassInstanceName(this MoveEndpoints endpoints)
        {
            return $"{endpoints.DestinationModuleName().ToLowerCaseFirstLetter()}1";
        }

        public static string DestinationModuleName(this MoveEndpoints endpoints)
        {
            return endpoints.ToDestinationAttributes().ModuleName;
        }

        public static ComponentType SourceComponentType(this MoveEndpoints endpoints)
        {
            return endpoints.ToDestinationAttributes().ComponentType;
        }

        public static ComponentType DestinationComponentType(this MoveEndpoints endpoints)
        {
            return endpoints.ToDestinationAttributes().ComponentType;
        }

        public static bool IsClassSource(this MoveEndpoints endpoints)
        {
            return endpoints == MoveEndpoints.ClassToClass || endpoints == MoveEndpoints.ClassToStd;
        }

        public static bool IsFormSource(this MoveEndpoints endpoints)
        {
            return endpoints == MoveEndpoints.FormToClass || endpoints == MoveEndpoints.FormToStd;
        }

        public static bool IsStdModuleSource(this MoveEndpoints endpoints)
        {
            return endpoints == MoveEndpoints.StdToClass || endpoints == MoveEndpoints.StdToStd;
        }

        private static (string ModuleName, ComponentType ComponentType) ToDestinationAttributes(this MoveEndpoints endpoints)
        {
            switch (endpoints)
            {
                case MoveEndpoints.StdToStd:
                case MoveEndpoints.ClassToStd:
                case MoveEndpoints.FormToStd:
                    return (DEFAULT_DESTINATION_MODULE_NAME, ComponentType.StandardModule);
                case MoveEndpoints.StdToClass:
                case MoveEndpoints.ClassToClass:
                case MoveEndpoints.FormToClass:
                    return (DEFAULT_DESTINATION_CLASS_NAME, ComponentType.ClassModule);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        private static (string ModuleName, ComponentType ComponentType) ToSourceAttributes(this MoveEndpoints endpoints)
        {
            switch (endpoints)
            {
                case MoveEndpoints.FormToStd:
                case MoveEndpoints.FormToClass:
                    return (DEFAULT_SOURCE_FORM_NAME, ComponentType.UserForm);
                case MoveEndpoints.ClassToStd:
                case MoveEndpoints.ClassToClass:
                    return (DEFAULT_SOURCE_CLASS_NAME, ComponentType.ClassModule);
                case MoveEndpoints.StdToStd:
                case MoveEndpoints.StdToClass:
                    return (DEFAULT_SOURCE_MODULE_NAME, ComponentType.StandardModule);
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }

        public static DeclarationType ToDeclarationType(this MoveEndpoints endpoints)
        {
            switch (endpoints)
            {
                case MoveEndpoints.StdToStd:
                case MoveEndpoints.ClassToStd:
                case MoveEndpoints.FormToStd:
                    return DeclarationType.ProceduralModule;
                case MoveEndpoints.StdToClass:
                case MoveEndpoints.ClassToClass:
                case MoveEndpoints.FormToClass:
                    return DeclarationType.ClassModule;
                default:
                    throw new ArgumentOutOfRangeException();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;
using Antlr4.Runtime;
using Rubberduck.UI;
using Rubberduck.Common;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public sealed class FunctionReturnValueNotUsedInspection : InspectionBase
    {
        public FunctionReturnValueNotUsedInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = CodeInspectionSeverity.Warning;
        }

        public override string Description { get { return InspectionsUI.FunctionReturnValueNotUsedInspection; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<CodeInspectionResultBase> GetInspectionResults()
        {
            // Note: This inspection does not find dictionary calls (e.g. foo!bar) since we do not know what the
            // default member is of a class.
            var interfaceMembers = UserDeclarations.FindInterfaceMembers();
            var interfaceImplementationMembers = UserDeclarations.FindInterfaceImplementationMembers();
            var functions = UserDeclarations.Where(function => function.DeclarationType == DeclarationType.Function).ToList();
            var interfaceMemberIssues = GetInterfaceMemberIssues(interfaceMembers);
            var nonInterfaceFunctions = functions.Except(interfaceMembers.Union(interfaceImplementationMembers));
            var nonInterfaceIssues = GetNonInterfaceIssues(nonInterfaceFunctions);
            return interfaceMemberIssues.Union(nonInterfaceIssues);
        }

        private IEnumerable<FunctionReturnValueNotUsedInspectionResult> GetInterfaceMemberIssues(IEnumerable<Declaration> interfaceMembers)
        {
            var interfaceIssues = new List<FunctionReturnValueNotUsedInspectionResult>();
            foreach (var interfaceMember in interfaceMembers)
            {
                var implementationMembers = UserDeclarations.FindInterfaceImplementationMembers(interfaceMember.IdentifierName);
                if (!IsReturnValueUsed(interfaceMember) && implementationMembers.All(member => !IsReturnValueUsed(member)))
                {
                    var implementationMemberIssues = implementationMembers
                        .Select(implementationMember =>
                            Tuple.Create(
                                implementationMember.Context,
                                new QualifiedSelection(
                                    implementationMember.QualifiedName.QualifiedModuleName,
                                    implementationMember.Selection),
                                GetReturnStatements(implementationMember)));

                    interfaceIssues.Add(new FunctionReturnValueNotUsedInspectionResult(
                            this,
                            interfaceMember.Context,
                            interfaceMember.QualifiedName,
                            GetReturnStatements(interfaceMember),
                            implementationMemberIssues));
                }
            }
            return interfaceIssues;
        }

        private IEnumerable<FunctionReturnValueNotUsedInspectionResult> GetNonInterfaceIssues(IEnumerable<Declaration> nonInterfaceFunctions)
        {
            var returnValueNotUsedFunctions = nonInterfaceFunctions.Where(function => !IsReturnValueUsed(function));
            var nonInterfaceIssues = returnValueNotUsedFunctions
                .Select(function =>
                        new FunctionReturnValueNotUsedInspectionResult(
                            this,
                            function.Context,
                            function.QualifiedName,
                            GetReturnStatements(function)));
            return nonInterfaceIssues;
        }

        private IEnumerable<string> GetReturnStatements(Declaration function)
        {
            return function.References
                .Where(usage => IsReturnStatement(function, usage))
                .Select(usage => usage.Context.Parent.Parent.Parent.GetText());
        }

        private bool IsReturnValueUsed(Declaration function)
        {
            return function.References.Count() > 0
                && function.References.Any(usage =>
                            !IsReturnStatement(function, usage) && !IsAddressOfCall(usage) && !IsCallWithoutAssignment(usage));
        }

        private bool IsAddressOfCall(IdentifierReference usage)
        {
            RuleContext current = usage.Context;
            while (current != null && !(current is VBAParser.VsAddressOfContext)) current = current.Parent;
            return current != null;
        }

        private bool IsReturnStatement(Declaration function, IdentifierReference assignment)
        {
            return assignment.ParentScope == function.Scope;
        }

        private bool IsCallWithoutAssignment(IdentifierReference usage)
        {
            return usage.Context.Parent != null && usage.Context.Parent.Parent is VBAParser.ImplicitCallStmt_InBlockContext;
        }
    }
}

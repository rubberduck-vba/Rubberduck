using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;
using Antlr4.Runtime;
using Rubberduck.Common;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections
{
    public sealed class FunctionReturnValueNotUsedInspection : InspectionBase
    {
        public FunctionReturnValueNotUsedInspection(RubberduckParserState state)
            : base(state)
        {
        }

        public override string Meta { get { return InspectionsUI.FunctionReturnValueNotUsedInspectionMeta; } }
        public override string Description { get { return InspectionsUI.FunctionReturnValueNotUsedInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<InspectionResultBase> GetInspectionResults()
        {
            // Note: This inspection does not find dictionary calls (e.g. foo!bar) since we do not know what the
            // default member is of a class.
            var interfaceMembers = UserDeclarations.FindInterfaceMembers().ToList();
            var interfaceImplementationMembers = UserDeclarations.FindInterfaceImplementationMembers();
            var functions = UserDeclarations.Where(function => function.DeclarationType == DeclarationType.Function).ToList();
            var interfaceMemberIssues = GetInterfaceMemberIssues(interfaceMembers);
            var nonInterfaceFunctions = functions.Except(interfaceMembers.Union(interfaceImplementationMembers));
            var nonInterfaceIssues = GetNonInterfaceIssues(nonInterfaceFunctions);
            return interfaceMemberIssues.Union(nonInterfaceIssues);
        }

        private IEnumerable<FunctionReturnValueNotUsedInspectionResult> GetInterfaceMemberIssues(IEnumerable<Declaration> interfaceMembers)
        {
            return from interfaceMember in interfaceMembers
                   let implementationMembers =
                       UserDeclarations.FindInterfaceImplementationMembers(interfaceMember.IdentifierName).ToList()
                   where
                       !IsReturnValueUsed(interfaceMember) &&
                       implementationMembers.All(member => !IsReturnValueUsed(member))
                   let implementationMemberIssues =
                       implementationMembers.Select(
                           implementationMember =>
                               Tuple.Create(implementationMember.Context,
                                   new QualifiedSelection(implementationMember.QualifiedName.QualifiedModuleName,
                                       implementationMember.Selection), GetReturnStatements(implementationMember)))
                   select
                       new FunctionReturnValueNotUsedInspectionResult(this, interfaceMember.Context,
                           interfaceMember.QualifiedName, GetReturnStatements(interfaceMember),
                           implementationMemberIssues, interfaceMember);
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
                            GetReturnStatements(function),
                            function));
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
            return function.References.Any(
                usage => 
                !IsReturnStatement(function, usage)
                && !IsAddressOfCall(usage)
                && !IsExplicitCall(usage)
                && !IsCallResultUsed(usage));
        }

        private bool IsAddressOfCall(IdentifierReference usage)
        {
            return ParserRuleContextHelper.HasParent<VBAParser.VsAddressOfContext>(usage.Context);
        }

        private bool IsReturnStatement(Declaration function, IdentifierReference assignment)
        {
            return assignment.ParentScoping.Equals(function);
        }

        private bool IsExplicitCall(IdentifierReference usage)
        {
            return ParserRuleContextHelper.HasParent<VBAParser.ExplicitCallStmtContext>(usage.Context);
        }

        private bool IsCallResultUsed(IdentifierReference usage)
        {
            return ParserRuleContextHelper.HasParent<VBAParser.ImplicitCallStmt_InBlockContext>(usage.Context)
                && !ParserRuleContextHelper.HasParent<VBAParser.ImplicitCallStmt_InStmtContext>(usage.Context)
                && !ParserRuleContextHelper.HasParent<VBAParser.LetStmtContext>(usage.Context)
                && !ParserRuleContextHelper.HasParent<VBAParser.SetStmtContext>(usage.Context);
        }
    }
}

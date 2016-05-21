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
            //// Temporarily disabled until fix for lack of context because of new resolver is found...
            //return new List<InspectionResultBase>();
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
            foreach (var usage in function.References)
            {
                if (IsReturnStatement(function, usage))
                {
                    continue;
                }
                if (IsAddressOfCall(usage))
                {
                    continue;
                }
                if (IsTypeOfExpression(usage))
                {
                    continue;
                }
                if (IsCallStmt(usage))
                {
                    continue;
                }
                if (IsLet(usage))
                {
                    continue;
                }
                if (IsSet(usage))
                {
                    continue;
                }
                return true;
            }
            return false;
        }

        private bool IsAddressOfCall(IdentifierReference usage)
        {
            var what = usage.Context.GetType();
            return ParserRuleContextHelper.HasParent<VBAParser.AddressOfExpressionContext>(usage.Context);
        }

        private bool IsTypeOfExpression(IdentifierReference usage)
        {
            return ParserRuleContextHelper.HasParent<VBAParser.TypeofexprContext>(usage.Context);
        }

        private bool IsReturnStatement(Declaration function, IdentifierReference assignment)
        {
            return assignment.ParentScoping.Equals(function) && assignment.Declaration.Equals(function);
        }

        private bool IsCallStmt(IdentifierReference usage)
        {
            var callStmt = ParserRuleContextHelper.GetParent<VBAParser.CallStmtContext>(usage.Context);
            if (callStmt == null)
            {
                return false;
            }
            var argumentList = CallStatement.GetArgumentList(callStmt);
            if (argumentList == null)
            {
                return true;
            }
            bool isUsedAsArgumentThusReturnValueIsUsed = ParserRuleContextHelper.HasParent(usage.Context, argumentList);
            return !isUsedAsArgumentThusReturnValueIsUsed;
        }

        private bool IsLet(IdentifierReference usage)
        {
            var letStmt = ParserRuleContextHelper.GetParent<VBAParser.LetStmtContext>(usage.Context);
            if (letStmt == null)
            {
                return false;
            }
            bool isLetAssignmentTarget = letStmt == usage.Context;
            return isLetAssignmentTarget;
        }

        private bool IsSet(IdentifierReference usage)
        {
            var setStmt = ParserRuleContextHelper.GetParent<VBAParser.SetStmtContext>(usage.Context);
            if (setStmt == null)
            {
                return false;
            }
            bool isSetAssignmentTarget = setStmt == usage.Context;
            return isSetAssignmentTarget;
        }
    }
}

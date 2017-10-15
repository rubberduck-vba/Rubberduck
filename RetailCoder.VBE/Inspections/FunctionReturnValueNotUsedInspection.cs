﻿using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
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
            var functions = State.DeclarationFinder
                .UserDeclarations(DeclarationType.Function)
                .Where(item => !IsIgnoringInspectionResultFor(item, AnnotationName))
                .ToList();
            var interfaceMemberIssues = GetInterfaceMemberIssues(interfaceMembers);
            var nonInterfaceFunctions = functions.Except(interfaceMembers.Union(interfaceImplementationMembers));
            var nonInterfaceIssues = GetNonInterfaceIssues(nonInterfaceFunctions);
            return interfaceMemberIssues.Union(nonInterfaceIssues);
        }

        private IEnumerable<FunctionReturnValueNotUsedInspectionResult> GetInterfaceMemberIssues(IEnumerable<Declaration> interfaceMembers)
        {
            return (from interfaceMember in interfaceMembers
                   let implementationMembers =
                       UserDeclarations.FindInterfaceImplementationMembers(interfaceMember.IdentifierName).ToList()
                   where interfaceMember.DeclarationType == DeclarationType.Function &&
                       !IsReturnValueUsed(interfaceMember) &&
                       implementationMembers.All(member => !IsReturnValueUsed(member))
                   let implementationMemberIssues =
                       implementationMembers.Select(
                           implementationMember =>
                               Tuple.Create(implementationMember.Context,
                                   new QualifiedSelection(implementationMember.QualifiedName.QualifiedModuleName,
                                       implementationMember.Selection), implementationMember))
                   select
                       new FunctionReturnValueNotUsedInspectionResult(this, interfaceMember.Context,
                           interfaceMember.QualifiedName, implementationMemberIssues, interfaceMember)).ToList();
        }

        private IEnumerable<FunctionReturnValueNotUsedInspectionResult> GetNonInterfaceIssues(IEnumerable<Declaration> nonInterfaceFunctions)
        {
            var returnValueNotUsedFunctions = nonInterfaceFunctions.Where(function => function.DeclarationType == DeclarationType.Function && !IsReturnValueUsed(function));
            var nonInterfaceIssues = returnValueNotUsedFunctions
                .Where(function => !IsRecursive(function))
                .Select(function =>
                        new FunctionReturnValueNotUsedInspectionResult(
                            this,
                            function.Context,
                            function.QualifiedName,
                            function));
            return nonInterfaceIssues;
        }

        private bool IsRecursive(Declaration function)
        {
            return function.References.Any(usage => usage.ParentScoping.Equals(function) && IsIndexExprOrCallStmt(usage));
        }

        private bool IsReturnValueUsed(Declaration function)
        {
            foreach (var usage in function.References)
            {
                if (IsAddressOfCall(usage))
                {
                    continue;
                }
                if (IsTypeOfExpression(usage))
                {
                    continue;
                }
                if (IsCallStmt(usage)) // IsIndexExprOrCallStmt(usage))
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
                if (IsReturnStatement(function, usage))
                {
                    continue;
                }
                return true;
            }
            return false;
        }

        private bool IsAddressOfCall(IdentifierReference usage)
        {
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

        private bool IsIndexExprOrCallStmt(IdentifierReference usage)
        {
            return IsCallStmt(usage) || IsIndexExprContext(usage);
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
            return !ParserRuleContextHelper.HasParent(usage.Context, argumentList);
        }

        private bool IsIndexExprContext(IdentifierReference usage)
        {
            var indexExpr = ParserRuleContextHelper.GetParent<VBAParser.IndexExprContext>(usage.Context);
            if (indexExpr == null)
            {
                return false;
            }
            var argumentList = indexExpr.argumentList();
            if (argumentList == null)
            {
                return true;
            }
            return !ParserRuleContextHelper.HasParent(usage.Context, argumentList);
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

using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.Inspections.Inspections.Extensions;

namespace Rubberduck.Inspections.Concrete
{
    /// <summary>
    /// Warns when a user function's return value is never used, at any of its call sites.
    /// </summary>
    /// <why>
    /// A 'Function' procedure normally means its return value to be captured and consumed by the calling code. 
    /// It's possible that not all call sites need the return value, but if the value is systematically discarded then this
    /// means the function is side-effecting, and thus should probably be a 'Sub' procedure instead.
    /// </why>
    /// <example hasResults="true">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     GetFoo ' return value is not captured
    /// End Sub
    /// 
    /// Private Function GetFoo() As Long
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    /// <example hasResults="false">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    ///     foo = GetFoo
    /// End Sub
    /// 
    /// Private Function GetFoo() As Long
    ///     GetFoo = 42
    /// End Function
    /// ]]>
    /// </example>
    public sealed class FunctionReturnValueNotUsedInspection : InspectionBase
    {
        public FunctionReturnValueNotUsedInspection(RubberduckParserState state)
            : base(state) { }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var interfaceMembers = State.DeclarationFinder.FindAllInterfaceMembers().ToList();
            var interfaceImplementationMembers = State.DeclarationFinder.FindAllInterfaceImplementingMembers();
            var functions = State.DeclarationFinder
                .UserDeclarations(DeclarationType.Function)
                .Where(item => item.References.Any(r => !IsReturnStatement(item, r) && !r.IsAssignment))
                .ToList();
            var interfaceMemberIssues = GetInterfaceMemberIssues(interfaceMembers);
            var nonInterfaceFunctions = functions.Except(interfaceMembers.Union(interfaceImplementationMembers));
            var nonInterfaceIssues = GetNonInterfaceIssues(nonInterfaceFunctions);
            return interfaceMemberIssues.Union(nonInterfaceIssues);
        }

        private IEnumerable<IInspectionResult> GetInterfaceMemberIssues(IEnumerable<Declaration> interfaceMembers)
        {
            return from interfaceMember in interfaceMembers
                let implementationMembers =
                    State.DeclarationFinder.FindInterfaceImplementationMembers(interfaceMember).ToList()
                where interfaceMember.DeclarationType == DeclarationType.Function &&
                      !IsReturnValueUsed(interfaceMember) &&
                      implementationMembers.All(member => !IsReturnValueUsed(member))
                let implementationMemberIssues =
                    implementationMembers.Select(
                        implementationMember =>
                            Tuple.Create(implementationMember.Context,
                                new QualifiedSelection(implementationMember.QualifiedName.QualifiedModuleName,
                                    implementationMember.Selection), implementationMember))
                select CreateInspectionResult(this, interfaceMember);

        }

        private IEnumerable<IInspectionResult> GetNonInterfaceIssues(IEnumerable<Declaration> nonInterfaceFunctions)
        {
            var returnValueNotUsedFunctions = nonInterfaceFunctions.Where(function => function.DeclarationType == DeclarationType.Function && !IsReturnValueUsed(function));
            var nonInterfaceIssues = returnValueNotUsedFunctions
                .Where(function => !IsRecursive(function))
                .Select(function =>
                        new DeclarationInspectionResult(
                            this,
                            string.Format(InspectionResults.FunctionReturnValueNotUsedInspection, function.IdentifierName),
                            function));
            return nonInterfaceIssues;
        }

        private bool IsRecursive(Declaration function)
        {
            return function.References.Any(usage => usage.ParentScoping.Equals(function) && IsIndexExprOrCallStmt(usage));
        }

        private bool IsReturnValueUsed(Declaration function)
        {
            // TODO: This is O(MG) at work here. Need to refactor the whole shebang.
            return (from usage in function.References
                where !IsLet(usage)
                where !IsSet(usage)
                where !IsCallStmt(usage)
                where !IsTypeOfExpression(usage)
                where !IsAddressOfCall(usage)
                select usage).Any(usage => !IsReturnStatement(function, usage));
        }

        private bool IsAddressOfCall(IdentifierReference usage)
        {
            return usage.Context.IsDescendentOf<VBAParser.AddressOfExpressionContext>();
        }

        private bool IsTypeOfExpression(IdentifierReference usage)
        {
            return usage.Context.IsDescendentOf<VBAParser.TypeofexprContext>();
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
            var callStmt = usage.Context.GetAncestor<VBAParser.CallStmtContext>();
            if (callStmt == null)
            {
                return false;
            }

            var indexExpr = usage.Context.GetAncestor<VBAParser.IndexExprContext>();
            if (indexExpr != null)
            {
                var memberAccessStmt = usage.Context.GetAncestor<VBAParser.MemberAccessExprContext>();
                if (memberAccessStmt != null &&
                    callStmt.SourceInterval.ProperlyContains(memberAccessStmt.SourceInterval) &&
                    memberAccessStmt.SourceInterval.ProperlyContains(indexExpr.SourceInterval))
                {
                    return false;
                }
            }

            var argumentList = CallStatement.GetArgumentList(callStmt);
            if (argumentList == null)
            {
                return true;
            }
            return !usage.Context.IsDescendentOf(argumentList);
        }

        private bool IsIndexExprContext(IdentifierReference usage)
        {
            var indexExpr = usage.Context.GetAncestor<VBAParser.IndexExprContext>();
            if (indexExpr == null)
            {
                return false;
            }
            var argumentList = indexExpr.argumentList();
            if (argumentList == null)
            {
                return true;
            }
            return !usage.Context.IsDescendentOf(argumentList);
        }

        private bool IsLet(IdentifierReference usage)
        {
            var letStmt = usage.Context.GetAncestor<VBAParser.LetStmtContext>();

            return letStmt != null && letStmt == usage.Context;
        }

        private bool IsSet(IdentifierReference usage)
        {
            var setStmt = usage.Context.GetAncestor<VBAParser.SetStmtContext>();

            return setStmt != null && setStmt == usage.Context;
        }

        private DeclarationInspectionResult CreateInspectionResult(IInspection inspection, Declaration interfaceMember)
        {
            dynamic properties = new PropertyBag();
            properties.DisableFixes = nameof(QuickFixes.ConvertToProcedureQuickFix);

            return new DeclarationInspectionResult(inspection,
                string.Format(InspectionResults.FunctionReturnValueNotUsedInspection,
                    interfaceMember.IdentifierName),
                interfaceMember, properties: properties);
        }
    }
}

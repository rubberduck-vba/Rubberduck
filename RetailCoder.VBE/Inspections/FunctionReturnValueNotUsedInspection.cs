using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;
using Antlr4.Runtime;
using Rubberduck.UI;

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
            var functions = UserDeclarations.Where(function => function.DeclarationType == DeclarationType.Function).ToList();
            var returnValueNotUsedFunctions = functions.Where(function =>
                    function.References.All(usage =>
                            IsReturnStatement(function, usage) || IsAddressOfCall(usage) || IsCallWithoutAssignment(usage)));

            var issues = returnValueNotUsedFunctions
                .Select(function =>
                        new FunctionReturnValueNotUsedInspectionResult(
                            this,
                            function.Context, 
                            function.QualifiedName,
                            function.References.Where(usage => IsReturnStatement(function, usage)).Select(usage => usage.Context.Parent.Parent.Parent.GetText())));

            return issues;
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
            return usage.Context.Parent != null && usage.Context.Parent is VBAParser.ICS_B_ProcedureCallContext;
        }
    }
}

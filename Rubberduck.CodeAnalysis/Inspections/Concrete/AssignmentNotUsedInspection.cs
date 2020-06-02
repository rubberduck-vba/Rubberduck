using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.Inspections.CodePathAnalysis;
using Rubberduck.Inspections.CodePathAnalysis.Extensions;
using Rubberduck.Inspections.CodePathAnalysis.Nodes;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Warns about a local variable that is assigned and never read. 
    /// Or, warns about a local variable that is assigned and then re-assigned 
    /// before using the previous value.
    /// </summary>
    /// <why>
    /// An assignment that is never used is a meaningless statement that was likely 
    /// used by an execution path in a prior version of the module.
    /// </why>
    /// <example hasResult="true">
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByRef value As Long)
    ///     Dim localVar As Long
    ///     localVar = 12 ' assignment never used
    ///     Dim otherVar As Long
    ///     otherVar = 12
    ///     value = otherVar * value
    /// End Sub
    /// ]]>
    /// <example hasResult="true">
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim localVar As Long
    ///     localVar = 12 ' assignment is redundant
    ///     localVar = 34 
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Function DoSomething(ByVal value As Long) As Long
    ///     Dim localVar As Long
    ///     localVar = 12
    ///     localVar = localVar + value 'variable is re-assigned, but the prior assigned value is read at least once first.
    ///     DoSomething = localVar
    /// End Function
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class AssignmentNotUsedInspection : IdentifierReferenceInspectionBase
    {
        private readonly Walker _walker;

        public AssignmentNotUsedInspection(IDeclarationFinderProvider declarationFinderProvider, Walker walker)
            : base(declarationFinderProvider)
        {
            _walker = walker;
        }

        protected override IEnumerable<IdentifierReference> ReferencesInModule(QualifiedModuleName module, DeclarationFinder finder)
        {
            var localNonArrayVariables = finder.Members(module, DeclarationType.Variable)
                .Where(declaration => !declaration.IsArray
                                      && !declaration.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Module));

            return localNonArrayVariables
                .Where(declaration => !declaration.IsIgnoringInspectionResultFor(AnnotationName))
                .SelectMany(d => FindUnusedAssignmentReferences(d, _walker));
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return !(IsAssignmentOfNothing(reference)
                        || IsPotentiallyUsedViaJump(reference, finder));
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            return Description;
        }

        private static IEnumerable<IdentifierReference> FindUnusedAssignmentReferences(Declaration localVariable, Walker walker)
        {
            if (!localVariable.References.Any(rf => rf.IsAssignment))
            {
                return Enumerable.Empty<IdentifierReference>();
            }

            //Consider static local variables used if they are referenced anywhere within their procedure
            if (localVariable.References.Any(r => !r.IsAssignment) && IsStatic(localVariable))
            {
                return Enumerable.Empty<IdentifierReference>();
            }

            var tree = walker.GenerateTree(localVariable.ParentScopeDeclaration.Context, localVariable);

            var allAssignmentsAndReferences = tree.Nodes(new[] { typeof(AssignmentNode), typeof(ReferenceNode) })
                                    .Where(node => localVariable.References.Contains(node.Reference));

            var unusedAssignmentNodes = allAssignmentsAndReferences.Any(n => n is ReferenceNode)
                        ? FindUnusedAssignmentNodes(tree, localVariable, allAssignmentsAndReferences)
                        : allAssignmentsAndReferences.OfType<AssignmentNode>();

            return unusedAssignmentNodes.Except(FindDescendantsOfNeverFlagNodeTypes(unusedAssignmentNodes))
                                        .Select(n => n.Reference);
        }

        private static IEnumerable<AssignmentNode> FindUnusedAssignmentNodes(INode node, Declaration localVariable, IEnumerable<INode> allAssignmentsAndReferences)
        {
            var assignmentExprNodes = node.Nodes(new[] { typeof(AssignmentExpressionNode) })
                                                .Where(n => localVariable.References.Contains(n.Children.FirstOrDefault()?.Reference));

            var usedAssignments = new List<AssignmentNode>();
            foreach (var refNode in allAssignmentsAndReferences.OfType<ReferenceNode>().Reverse())
            {
                var assignmentExprNodesWithReference = assignmentExprNodes.Where(n => n.Nodes(new[] { typeof(ReferenceNode) })
                                                            .Contains(refNode));

                var assignmentsPrecedingReference = assignmentExprNodesWithReference.Any()
                    ? assignmentExprNodes.TakeWhile(n => n != assignmentExprNodesWithReference.Last())
                                                .Last()
                                                .Nodes(new[] { typeof(AssignmentNode) })
                    : allAssignmentsAndReferences.TakeWhile(n => n != refNode)
                        .OfType<AssignmentNode>();

                if (assignmentsPrecedingReference.Any())
                {
                    usedAssignments.Add(assignmentsPrecedingReference.Last() as AssignmentNode);
                }
            }

            return allAssignmentsAndReferences.OfType<AssignmentNode>()
                                                .Except(usedAssignments);
        }

        private static IEnumerable<AssignmentNode> FindDescendantsOfNeverFlagNodeTypes(IEnumerable<AssignmentNode> flaggedAssignments)
        {
            var filteredResults = new List<AssignmentNode>();

            foreach (var assignment in flaggedAssignments)
            {
                if (assignment.TryGetAncestorNode<BranchNode>(out _))
                {
                    filteredResults.Add(assignment);
                }
                if (assignment.TryGetAncestorNode<LoopNode>(out _))
                {
                    filteredResults.Add(assignment);
                }
            }
            return filteredResults;
        }

        private static bool IsAssignmentOfNothing(IdentifierReference reference)
        {
            if (reference.Context.Parent is VBAParser.SetStmtContext setStmtContext2)
            {
                var test = setStmtContext2.expression();
            }
            return reference.IsSetAssignment
                && reference.Context.Parent is VBAParser.SetStmtContext setStmtContext
                && setStmtContext.expression().GetText().Equals(Tokens.Nothing);
        }

        /// <summary>
        /// Filters false positive result references due to GoTo and Resume statements.  e.g., 
        /// An ErrorHandler block that branches execution to a location where the asignment may be used. 
        /// </summary>
        /// <remarks>
        /// Filters Assignment references that meet the following conditions:
        /// 1. Precedes a GoTo or Resume statement that branches execution to a line before the 
        ///     assignment reference, and
        /// 2. A non-assignment reference is present on a line that is:
        ///     a) At or below the start of the execution branch, and 
        ///     b) Above the next ExitStatement line (if one exists) or the end of the procedure
        /// </remarks>
        private static bool IsPotentiallyUsedViaJump(IdentifierReference resultCandidate, DeclarationFinder finder)
        {
            if (!resultCandidate.Declaration.References.Any(rf => !rf.IsAssignment)) { return false; }

            var labelIdLineNumberPairs = finder.DeclarationsWithType(DeclarationType.LineLabel)
                                                .Where(label => resultCandidate.ParentScoping.Equals(label.ParentDeclaration))
                                                .Select(lbl => (lbl.IdentifierName, lbl.Context.Start.Line));

            return JumpStmtPotentiallyUsesVariable<VBAParser.GoToStmtContext>(resultCandidate, labelIdLineNumberPairs)
                || JumpStmtPotentiallyUsesVariable<VBAParser.ResumeStmtContext>(resultCandidate, labelIdLineNumberPairs);
        }

        private static bool JumpStmtPotentiallyUsesVariable<T>(IdentifierReference resultCandidate, IEnumerable<(string IdentifierName, int Line)> labelIdLineNumberPairs) where T : ParserRuleContext
        {
            if (TryGetRelevantJumpContext<T>(resultCandidate, out var jumpStmt))
            {
                return IsPotentiallyUsedAssignment(jumpStmt, resultCandidate, labelIdLineNumberPairs);
            }

            return false;
        }

        private static bool TryGetRelevantJumpContext<T>(IdentifierReference resultCandidate, out T ctxt) where T : ParserRuleContext
        {
            ctxt = resultCandidate.ParentScoping.Context.GetDescendents<T>()
                                    .Where(sc => sc.Start.Line > resultCandidate.Context.Start.Line
                                                    || (sc.Start.Line == resultCandidate.Context.Start.Line
                                                            && sc.Start.Column > resultCandidate.Context.Start.Column))
                                    .OrderBy(sc => sc.Start.Line - resultCandidate.Context.Start.Line)
                                    .ThenBy(sc => sc.Start.Column - resultCandidate.Context.Start.Column)
                                    .FirstOrDefault();
            return ctxt != null;
        }

        private static bool IsPotentiallyUsedAssignment<T>(T jumpContext, IdentifierReference resultCandidate, IEnumerable<(string, int)> labelIdLineNumberPairs)
        {
            int? executionBranchLine = null;
            if (jumpContext is VBAParser.GoToStmtContext gotoCtxt)
            {
                executionBranchLine = DetermineLabeledExecutionBranchLine(gotoCtxt.expression().GetText(), labelIdLineNumberPairs);
            }
            else
            {
                executionBranchLine = DetermineResumeStmtExecutionBranchLine(jumpContext as VBAParser.ResumeStmtContext, resultCandidate, labelIdLineNumberPairs);
            }

            return executionBranchLine.HasValue
                ?   AssignmentIsUsedPriorToExitStmts(resultCandidate, executionBranchLine.Value)
                :   false;
        }

        private static bool AssignmentIsUsedPriorToExitStmts(IdentifierReference resultCandidate, int executionBranchLine)
        {
            if (resultCandidate.Context.Start.Line < executionBranchLine) { return false; }

            var procedureExitStmtCtxts = resultCandidate.ParentScoping.Context.GetDescendents<VBAParser.ExitStmtContext>()
                                    .Where(exitCtxt => exitCtxt.EXIT_DO() == null
                                            && exitCtxt.EXIT_FOR() == null);

            var nonAssignmentCtxts = resultCandidate.Declaration.References
                                            .Where(rf => !rf.IsAssignment)
                                            .Select(rf => rf.Context);

            var sortedContextsAfterBranch = nonAssignmentCtxts.Concat(procedureExitStmtCtxts)
                        .Where(ctxt => ctxt.Start.Line >= executionBranchLine)
                        .OrderBy(ctxt => ctxt.Start.Line)
                        .ThenBy(ctxt => ctxt.Start.Column);

            return !(sortedContextsAfterBranch.FirstOrDefault() is VBAParser.ExitStmtContext);
        }

        private static int? DetermineResumeStmtExecutionBranchLine(VBAParser.ResumeStmtContext resumeStmt, IdentifierReference resultCandidate, IEnumerable<(string IdentifierName, int Line)> labelIdLineNumberPairs)
        {
            var onErrorGotoLabelToLineNumber = resultCandidate.ParentScoping.Context.GetDescendents<VBAParser.OnErrorStmtContext>()
                    .Where(errorStmtCtxt => !errorStmtCtxt.expression().GetText().Equals("0"))
                    .ToDictionary(k => k.expression()?.GetText() ?? "No Label", v => v.Start.Line);

            var errorHandlerLabelsAndLines = labelIdLineNumberPairs
                                                    .Where(pair => onErrorGotoLabelToLineNumber.ContainsKey(pair.IdentifierName));

            //Labels must be located at the start of a line.
            //If the resultCandidate line precedes all error handling related labels, 
            //a Resume statement cannot be invoked (successfully) for the resultCandidate
            if (!errorHandlerLabelsAndLines.Any(s => s.Line <= resultCandidate.Context.Start.Line))
            {
                return null;
            }

            var expression = resumeStmt.expression()?.GetText();

            //For Resume and Resume Next, expression() is null
            if (string.IsNullOrEmpty(expression))
            {
                //Get errorHandlerLabel for the Resume statement
                string errorHandlerLabel = errorHandlerLabelsAndLines
                                                .Where(pair => resumeStmt.Start.Line >= pair.Line)
                                                .OrderBy(pair => resumeStmt.Start.Line - pair.Line)
                                                .Select(pair => pair.IdentifierName)
                                                .FirstOrDefault();

                //Since the execution branch line for Resume and Resume Next statements 
                //is indeterminant by static analysis, the On***GoTo statement
                //is used as the execution branch line
                return onErrorGotoLabelToLineNumber[errorHandlerLabel];
            }
            //Resume <label>
            return DetermineLabeledExecutionBranchLine(expression, labelIdLineNumberPairs);
        }

        private static int DetermineLabeledExecutionBranchLine(string expression, IEnumerable<(string IdentifierName, int Line)> IDandLinePairs)
                        => int.TryParse(expression, out var parsedLineNumber)
                                        ? parsedLineNumber
                                        : IDandLinePairs.Single(v => v.IdentifierName.Equals(expression)).Line;

        private static bool IsStatic(Declaration declaration)
        {
            var ctxt = declaration.Context.GetAncestor<VBAParser.VariableStmtContext>();
            if (ctxt?.STATIC() != null)
            {
                return true;
            }

            switch (declaration.ParentDeclaration.Context)
            {
                case VBAParser.FunctionStmtContext func:
                    return func.STATIC() != null;
                case VBAParser.SubStmtContext sub:
                    return sub.STATIC() != null;
                case VBAParser.PropertyLetStmtContext let:
                    return let.STATIC() != null;
                case VBAParser.PropertySetStmtContext set:
                    return set.STATIC() != null;
                case VBAParser.PropertyGetStmtContext get:
                    return get.STATIC() != null;
                default:
                    return false;
            }
        }
    }
}

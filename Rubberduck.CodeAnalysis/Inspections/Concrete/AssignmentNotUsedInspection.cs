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
    /// </module>
    /// </example>
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
            return !(IsAssignmentOfNothing(reference) || IsPotentiallyUsedViaJump(reference, finder));
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
                                    .Where(node => localVariable.References.Contains(node.Reference))
                                    .ToList();

            var unusedAssignmentNodes = allAssignmentsAndReferences.Any(n => n is ReferenceNode)
                        ? FindUnusedAssignmentNodes(tree, localVariable, allAssignmentsAndReferences)
                        : allAssignmentsAndReferences.OfType<AssignmentNode>();

            var results = unusedAssignmentNodes
                .Where(n => !IsDescendentOfNeverFlagNode(n))
                .Select(n => n.Reference);

            return results;
        }

        private static IEnumerable<AssignmentNode> FindUnusedAssignmentNodes(INode node, Declaration localVariable, IEnumerable<INode> allAssignmentsAndReferences)
        {
            var assignmentExprNodes = node.Nodes(new[] { typeof(AssignmentExpressionNode) })
                                                .Where(n => localVariable.References.Contains(n.Children.FirstOrDefault()?.Reference))
                                                .ToList();

            var usedAssignments = new List<AssignmentNode>();
            foreach (var refNode in allAssignmentsAndReferences.OfType<ReferenceNode>().Reverse())
            {
                var assignmentExprNodesWithReference = assignmentExprNodes
                                                            .Where(n => n.Nodes(new[] { typeof(ReferenceNode) })
                                                            .Contains(refNode));

                var assignmentsPrecedingReference = assignmentExprNodesWithReference.Any()
                    ? assignmentExprNodes.TakeWhile(n => n != assignmentExprNodesWithReference.LastOrDefault())
                                                ?.LastOrDefault()
                                                ?.Nodes(new[] { typeof(AssignmentNode) })
                    : allAssignmentsAndReferences.TakeWhile(n => n != refNode && !IsDescendentOfNeverFlagNode(n))
                        .OfType<AssignmentNode>();

                if (assignmentsPrecedingReference?.Any() ?? false)
                {
                    usedAssignments.Add(assignmentsPrecedingReference.LastOrDefault() as AssignmentNode);
                }
            }

            return allAssignmentsAndReferences.OfType<AssignmentNode>().Except(usedAssignments);
        }

        private static bool IsDescendentOfNeverFlagNode(INode assignment)
        {
            return assignment.TryGetAncestorNode<BranchNode>(out _)
                || assignment.TryGetAncestorNode<LoopNode>(out _);
        }

        private static bool IsAssignmentOfNothing(IdentifierReference reference)
        {
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
        /// 1. Reference precedes a GoTo or Resume statement that branches execution to a line before the 
        ///     assignment reference, AND
        /// 2. A non-assignment reference is present on a line that is:
        ///     a) At or below the start of the execution branch, AND 
        ///     b) Above the next ExitStatement line (if one exists) or the end of the procedure
        /// </remarks>
        private static bool IsPotentiallyUsedViaJump(IdentifierReference resultCandidate, DeclarationFinder finder)
        {
            if (!resultCandidate.Declaration.References.Any(rf => !rf.IsAssignment)) { return false; }

            var labelIdLineNumberPairs = finder.Members(resultCandidate.QualifiedModuleName, DeclarationType.LineLabel)
                                                .Where(label => resultCandidate.ParentScoping.Equals(label.ParentDeclaration))
                                                .ToDictionary(key => key.IdentifierName, v => v.Context.Start.Line);

            return JumpStmtPotentiallyUsesVariable<VBAParser.GoToStmtContext>(resultCandidate, labelIdLineNumberPairs)
                || JumpStmtPotentiallyUsesVariable<VBAParser.ResumeStmtContext>(resultCandidate, labelIdLineNumberPairs);
        }

        private static bool JumpStmtPotentiallyUsesVariable<T>(IdentifierReference resultCandidate, Dictionary<string,int> labelIdLineNumberPairs) where T: ParserRuleContext
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
                                    .Where(descendent => descendent.GetSelection() > resultCandidate.Selection)
                                    .OrderBy(descendent => descendent.GetSelection())
                                    .FirstOrDefault();
            return ctxt != null;
        }

        private static bool IsPotentiallyUsedAssignment<T>(T jumpContext, IdentifierReference resultCandidate, Dictionary<string, int> labelIdLineNumberPairs) where T : ParserRuleContext
        {
            int? executionBranchLine;

            switch (jumpContext)
            {
                case VBAParser.GoToStmtContext gotoStmt:
                    executionBranchLine = labelIdLineNumberPairs[gotoStmt.expression().GetText()];
                    break;
                case VBAParser.ResumeStmtContext resume:
                    executionBranchLine = DetermineResumeStmtExecutionBranchLine(resume, resultCandidate, labelIdLineNumberPairs);
                    break;
                default:
                    executionBranchLine = null;
                    break;
            }

            return executionBranchLine.HasValue && AssignmentIsUsedPriorToExitStmts(resultCandidate, executionBranchLine.Value);
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

        private static int? DetermineResumeStmtExecutionBranchLine(VBAParser.ResumeStmtContext resumeStmt, IdentifierReference resultCandidate, Dictionary<string, int> labelIdLineNumberPairs)
        {
            var onErrorGotoLabelToLineNumber = resultCandidate.ParentScoping.Context.GetDescendents<VBAParser.OnErrorStmtContext>()
                    .Where(errorStmtCtxt => IsBranchingOnErrorGoToLabel(errorStmtCtxt))
                    .ToDictionary(k => k.expression()?.GetText() ?? "No Label", v => v.Start.Line);

            var errorHandlerLabelsAndLines = labelIdLineNumberPairs
                                                    .Where(pair => onErrorGotoLabelToLineNumber.ContainsKey(pair.Key));

            //Labels must be located at the start of a line.
            //If the resultCandidate line precedes all error handling related labels, 
            //a Resume statement cannot be invoked (successfully) for the resultCandidate
            if (!errorHandlerLabelsAndLines.Any(kvp => kvp.Value <= resultCandidate.Context.Start.Line))
            {
                return null;
            }

            var resumeStmtExpression = resumeStmt.expression()?.GetText();

            //For Resume and Resume Next, expression() is null
            if (string.IsNullOrEmpty(resumeStmtExpression))
            {
                var errorHandlerLabelForTheResumeStatement = errorHandlerLabelsAndLines
                                                .Where(kvp => resumeStmt.Start.Line >= kvp.Value)
                                                .OrderBy(kvp => resumeStmt.Start.Line - kvp.Value)
                                                .Select(kvp => kvp.Key)
                                                .FirstOrDefault();

                //Since the execution branch line for Resume and Resume Next statements 
                //is indeterminant by static analysis, the On***GoTo statement
                //is used as the execution branch line
                return onErrorGotoLabelToLineNumber[errorHandlerLabelForTheResumeStatement];
            }
            //Resume <label>
            return labelIdLineNumberPairs[resumeStmtExpression];
        }

        private static bool IsBranchingOnErrorGoToLabel(VBAParser.OnErrorStmtContext errorStmtCtxt)
        {
            var label = errorStmtCtxt.expression()?.GetText();
            if (string.IsNullOrEmpty(label))
            {
                return false;
            }
            //The VBE will complain about labels other than:
            //1. Numerics less than int.MaxValue (VBA: 'Long' max value).  '0' returns false because it cause a branch
            //2. Or, Any alphanumeric string beginning with a letter (VBE or the Debugger will choke on special characters, spaces, etc).
            return !(int.TryParse(label, out var numberLabel) && numberLabel <= 0);
        }

        //TODO: Add IsStatic member to VariableDeclaration
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

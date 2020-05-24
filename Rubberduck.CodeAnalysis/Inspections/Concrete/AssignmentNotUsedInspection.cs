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
    /// Warns about a variable that is assigned, and then re-assigned before the first assignment is read.
    /// </summary>
    /// <why>
    /// The first assignment is likely redundant, since it is being overwritten by the second.
    /// </why>
    /// <example hasResult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething()
    ///     Dim foo As Long
    ///     foo = 12 ' assignment is redundant
    ///     foo = 34 
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasResult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Sub DoSomething(ByVal foo As Long)
    ///     Dim bar As Long
    ///     bar = 12
    ///     bar = bar + foo ' variable is re-assigned, but the prior assigned value is read at least once first.
    /// End Sub
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
                .SelectMany(UnusedAssignments);
        }

        private IEnumerable<IdentifierReference> UnusedAssignments(Declaration localVariable)
        {
            var tree = _walker.GenerateTree(localVariable.ParentScopeDeclaration.Context, localVariable);
            return UnusedAssignmentReferences(tree);
        }

        private static List<IdentifierReference> UnusedAssignmentReferences(INode node)
        {
            var nodes = new List<IdentifierReference>();

            var blockNodes = node.Nodes(new[] { typeof(BlockNode) });
            foreach (var block in blockNodes)
            {
                INode lastNode = default;
                foreach (var flattenedNode in block.FlattenedNodes(new[] { typeof(GenericNode), typeof(BlockNode) }))
                {
                    if (flattenedNode is AssignmentNode &&
                        lastNode is AssignmentNode)
                    {
                        nodes.Add(lastNode.Reference);
                    }

                    lastNode = flattenedNode;
                }

                if (lastNode is AssignmentNode &&
                    block.Children[0].GetFirstNode(new[] { typeof(GenericNode) }) is DeclarationNode)
                {
                    nodes.Add(lastNode.Reference);
                }
            }

            return nodes;
        }

        protected override bool IsResultReference(IdentifierReference reference, DeclarationFinder finder)
        {
            return !(IsAssignmentOfNothing(reference)
                        || IsPotentiallyUsedViaJump(reference, finder));
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

            return GotoPotentiallyUsesVariable(resultCandidate, labelIdLineNumberPairs) 
                || ResumePotentiallyUsesVariable(resultCandidate, labelIdLineNumberPairs);
        }

        private static bool GotoPotentiallyUsesVariable(IdentifierReference resultCandidate, IEnumerable<(string, int)> labelIdLineNumberPairs)
        {
            if (TryGetRelevantJumpContext<VBAParser.GoToStmtContext>(resultCandidate, out var gotoStmt))
            {
                return IsPotentiallyUsedAssignment(gotoStmt, resultCandidate, labelIdLineNumberPairs);
            }

            return false;
        }

        private static bool ResumePotentiallyUsesVariable(IdentifierReference resultCandidate, IEnumerable<(string IdentifierName, int Line)> labelIdLineNumberPairs)
        {
            if (TryGetRelevantJumpContext<VBAParser.ResumeStmtContext>(resultCandidate, out var resumeStmt))
            {
                return IsPotentiallyUsedAssignment(resumeStmt, resultCandidate, labelIdLineNumberPairs);
            }

            return false;
        }

        private static bool TryGetRelevantJumpContext<T>(IdentifierReference resultCandidate, out T ctxt) where T : ParserRuleContext //, IEnumerable<T> stmtContexts, int targetLine, int? targetColumn = null) where T : ParserRuleContext
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

        private static bool IsPotentiallyUsedAssignment<T>(T jumpContext, IdentifierReference resultCandidate, IEnumerable<(string, int)> labelIdLineNumberPairs) //, int executionBranchLine)
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

        private static int? DetermineResumeStmtExecutionBranchLine(VBAParser.ResumeStmtContext resumeStmt, IdentifierReference resultCandidate, IEnumerable<(string IdentifierName, int Line)> labelIdLineNumberPairs) //where T: ParserRuleContext
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

        protected override string ResultDescription(IdentifierReference reference)
        {
            return Description;
        }
    }
}

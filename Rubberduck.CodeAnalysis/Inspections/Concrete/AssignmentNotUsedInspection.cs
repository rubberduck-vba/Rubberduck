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
                        || IsPotentiallyUsedViaResumeOrGoToExecutionBranch(reference, finder));
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
        /// Excludes Assignment references that meet the following conditions:
        /// 1. Preceed a GoTo or Resume statement that branches execution to a line before the 
        ///     assignment reference, and
        /// 2. A non-assignment reference is present on a line that is:
        ///     a) At or below the start of the execution branch, and 
        ///     b) Above the next ExitStatement line (if one exists) or the end of the procedure
        /// </remarks>
        /// <param name="resultCandidate"></param>
        /// <param name="finder"></param>
        /// <returns></returns>
        private static bool IsPotentiallyUsedViaResumeOrGoToExecutionBranch(IdentifierReference resultCandidate, DeclarationFinder finder)
        {
            if (!resultCandidate.Declaration.References.Any(rf => !rf.IsAssignment)) { return false; }

            var labelIdLineNumberPairs = finder.DeclarationsWithType(DeclarationType.LineLabel)
                                                .Where(label => resultCandidate.ParentScoping.Equals(label.ParentDeclaration))
                                                .Select(lbl => (lbl.IdentifierName, lbl.Context.Start.Line));

            return GotoExecutionBranchPotentiallyUsesVariable(resultCandidate, labelIdLineNumberPairs) 
                || ResumeExecutionBranchPotentiallyUsesVariable(resultCandidate, labelIdLineNumberPairs);
        }

        private static bool GotoExecutionBranchPotentiallyUsesVariable(IdentifierReference resultCandidate, IEnumerable<(string, int)> labelIdLineNumberPairs)
        {
            var gotoCtxts = resultCandidate.ParentScoping.Context.GetDescendents<VBAParser.GoToStmtContext>()
                .Where(gotoCtxt => gotoCtxt.Start.Line > resultCandidate.Context.Start.Line);

            if (!gotoCtxts.Any()) { return false; }

            var gotoStmt = GetFirstContextAfterLine(gotoCtxts, resultCandidate.Context.Start.Line);

            if (gotoStmt == null) { return false; }

            var executionBranchLine = DetermineExecutionBranchLine(gotoStmt.expression().GetText(), labelIdLineNumberPairs);

            return IsPotentiallyUsedAssignment(resultCandidate, executionBranchLine);
        }

        private static bool ResumeExecutionBranchPotentiallyUsesVariable(IdentifierReference resultCandidate, IEnumerable<(string IdentifierName, int Line)> labelIdLineNumberPairs)
        {
            var resumeStmtCtxts = resultCandidate.ParentScoping.Context.GetDescendents<VBAParser.ResumeStmtContext>()
                .Where(jumpCtxt => jumpCtxt.Start.Line > resultCandidate.Context.Start.Line);

            if (!resumeStmtCtxts.Any()) { return false; }

            var onErrorGotoStatements = resultCandidate.ParentScoping.Context.GetDescendents<VBAParser.OnErrorStmtContext>()
                                .Where(errorStmtCtxt => !errorStmtCtxt.expression().GetText().Equals("0"))
                                .ToDictionary(k => k.expression()?.GetText() ?? "0", v => v.Start.Line);

            var errorHandlerLabelsAndLines = labelIdLineNumberPairs.Where(pair => onErrorGotoStatements.ContainsKey(pair.IdentifierName));
            
            //If the resultCandidate line preceeds all ErrorHandlers/Resume statements - it is not evaluated
            if (errorHandlerLabelsAndLines.All(s => s.Line > resultCandidate.Context.Start.Line))
            {
                return false;
            }

            var resumeStmt = GetFirstContextAfterLine(resumeStmtCtxts, resultCandidate.Context.Start.Line);

            if (resumeStmt == null) { return false; }

            int? executionBranchLine = null;

            var expression = resumeStmt.expression()?.GetText();

            //For Resume and Resume Next, expression() is null
            if (string.IsNullOrEmpty(expression))
            {
                //Get info for the errorHandlerLabel above the Resume statement
                (string IdentifierName, int Line)? errorHandlerLabel = labelIdLineNumberPairs
                                                            .Where(pair => resumeStmt.Start.Line > pair.Line)
                                                            .OrderBy(pair => resumeStmt.Start.Line - pair.Line)
                                                            .FirstOrDefault();

                //Since the execution branch line for Resume and Resume Next statements 
                //is indeterminant by static analysis, the On***GoTo statement
                //is used as the execution branch line
                if (errorHandlerLabel.HasValue && onErrorGotoStatements.ContainsKey(errorHandlerLabel.Value.IdentifierName))
                {
                    executionBranchLine = onErrorGotoStatements[errorHandlerLabel.Value.IdentifierName];
                }
            }
            else
            {
                executionBranchLine = DetermineExecutionBranchLine(expression, labelIdLineNumberPairs);
            }

            return executionBranchLine.HasValue
                ? IsPotentiallyUsedAssignment(resultCandidate, executionBranchLine.Value)
                : false;
        }

        private static bool IsPotentiallyUsedAssignment(IdentifierReference resultCandidate, int executionBranchLine)
        {
            if (resultCandidate.Context.Start.Line <= executionBranchLine) { return false; }

            var exitStmtCtxts = resultCandidate.ParentScoping.Context.GetDescendents<VBAParser.ExitStmtContext>()
                                .Where(exitCtxt => exitCtxt.Start.Line > executionBranchLine
                                        && exitCtxt.EXIT_DO() == null
                                        && exitCtxt.EXIT_FOR() == null);

            var exitStmtCtxt = GetFirstContextAfterLine(exitStmtCtxts, executionBranchLine);

            var nonAssignmentReferences = resultCandidate.Declaration.References
                                                .Where(rf => !rf.IsAssignment);

            var possibleUse = exitStmtCtxt != null
                ? nonAssignmentReferences.Where(rf => rf.Context.Start.Line >= executionBranchLine
                                                        && rf.Context.Start.Line < exitStmtCtxt.Start.Line)
                : nonAssignmentReferences.Where(rf => rf.Context.Start.Line >= executionBranchLine);

            return possibleUse.Any();
        }

        private static int DetermineExecutionBranchLine(string expression, IEnumerable<(string IdentifierName, int Line)> IDandLinePairs)
        {
            if (int.TryParse(expression, out var parsedLineNumber))
            {
                return parsedLineNumber;
            }
            (string label, int lineNumber) = IDandLinePairs.Where(v => v.IdentifierName.Equals(expression)).Single();
            return lineNumber;
        }

        private static T GetFirstContextAfterLine<T>(IEnumerable<T> stmtContexts, int targetLine) where T : ParserRuleContext 
                                        => stmtContexts.Where(sc => sc.Start.Line > targetLine)
                                                        .OrderBy(sc => sc.Start.Line - targetLine)
                                                        .FirstOrDefault();

        protected override string ResultDescription(IdentifierReference reference)
        {
            return Description;
        }
    }
}

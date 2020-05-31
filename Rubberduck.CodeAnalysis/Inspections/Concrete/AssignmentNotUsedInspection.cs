using System.Collections.Generic;
using System.IO;
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
    /// <module name="Module1" type="Standard Module">
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
    /// <module name="Module1" type="Standard Module">
    /// <![CDATA[
    /// Public Function DoSomething(ByVal foo As Long) As Long
    ///     Dim bar As Long
    ///     bar = 12
    ///     bar = bar + foo ' variable is re-assigned, but the prior assigned value is read at least once first.
    ///     DoSomething = bar
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
                .SelectMany(UnusedAssignments);
        }

        private IEnumerable<IdentifierReference> UnusedAssignments(Declaration localVariable)
        {
            if (!localVariable.References.Any(rf => rf.IsAssignment))
            {
                return Enumerable.Empty<IdentifierReference>();
            }

            var tree = _walker.GenerateTree(localVariable.ParentScopeDeclaration.Context, localVariable);
            if (File.Exists("C:\\Users\\Brian\\Documents\\GitHub\\CodePath1.txt"))
            {
                return UnusedAssignmentReferences1(tree, localVariable);
            }
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


        public static IEnumerable<IdentifierReference> UnusedAssignmentReferences1(INode node, Declaration localVariable)
        {
            var allNodes = node.Nodes(new[] { typeof(AssignmentNode), typeof(ReferenceNode) })
                                    .Where(n => localVariable.References.Contains(n.Reference));

            var unUsedNodes = new List<AssignmentNode>();

            if (!allNodes.OfType<ReferenceNode>().Any())
            {
                unUsedNodes.AddRange(allNodes.OfType<AssignmentNode>().Cast<AssignmentNode>());
            }
            else
            {
                unUsedNodes.AddRange(AssignmentsTrailingLastReference(allNodes));
                unUsedNodes.AddRange(UnusedAssignments(allNodes));
            }

            return unUsedNodes.Except(DescendantsOfNonEvaluatedTypes(allNodes))
                                .Select(n => n.Reference);
        }

        /// <summary>
        /// Returns Assignments occuring after the last ReferenceNode
        /// </summary>
        /// <example>
        /// <code>
        ///Private Function Message()
        ///      fizz = "Hello"
        ///      Message = fizz
        ///      fizz = "GoodBye"   'Not used   
        ///      fizz = "Ciao"      'Not used  
        ///End Function
        /// </code>
        /// </example>
        private static IEnumerable<AssignmentNode> AssignmentsTrailingLastReference(IEnumerable<INode> allNodes)
        {
            var lastReferenceNode = allNodes.OfType<ReferenceNode>().Last();
            return allNodes
                        .SkipWhile(n => n != lastReferenceNode)
                        .OfType<AssignmentNode>();
        }
        /// <summary>
        /// Detects AssignmentNodes that are subsequently re-assigned without being referenced.  
        /// </summary>
        /// <remarks>
        /// There is a scenario where the sequential assignment analysis would generate 
        /// a false-positive result.  The scenario is addressed by this function.  See example 
        /// </remarks>
        /// <example>
        /// <code>
        /// Private Sub Test(ByRef value As Long)
        ///     Dim fizz As Long
        ///     fizz = value        'not used
        ///     fizz = value + 2    'not used
        ///     'The next statement 'looks' unused by sequential analysis only.
        ///     'But, it the assignment is referenced by the subsequent assignment
        ///     fizz = 6            
        ///     fizz = value + fizz  'not used
        ///     fizz = 7
        ///     value = fizz
        /// End Sub
        /// </code>
        /// </example>
        private static IEnumerable<AssignmentNode> UnusedAssignments(IEnumerable<INode> allNodes)
        {
            var assignmentsByReferenceNode = new Dictionary<ReferenceNode, List<AssignmentNode>>();
            foreach (var rfn in allNodes.OfType<ReferenceNode>())
            {
                assignmentsByReferenceNode.Add(rfn, allNodes.TakeWhile(n => n != rfn)
                                                        .OfType<AssignmentNode>()
                                                        .Except(assignmentsByReferenceNode.Values.SelectMany(v => v))
                                                        .Cast<AssignmentNode>().ToList());
            }

            var unUsedSequentialAssignments = new List<AssignmentNode>();

            var usedByNextAssignment = new List<AssignmentNode>();

            var sequentialAssignmentsFound = false;
            foreach (var key in assignmentsByReferenceNode.Keys)
            {
                if (assignmentsByReferenceNode[key].Count() > 1)
                {
                    sequentialAssignmentsFound = true;
                    var numberOfUnusedAssignments = assignmentsByReferenceNode[key].Count() - 1;
                    unUsedSequentialAssignments.AddRange(assignmentsByReferenceNode[key].Take(numberOfUnusedAssignments));
                }

                //Find assignments that are referenced in the RHS of the 'next' assignment
                if (unUsedSequentialAssignments.Any() && TryFindAssigmentUsedBySubsequentAssignment(key, assignmentsByReferenceNode[key], out var usedAssignment))
                {
                    usedByNextAssignment.Add(usedAssignment);
                }
            }

            if (sequentialAssignmentsFound)
            {
                var lastAssignmentNode = allNodes.OfType<AssignmentNode>().Last();

                if (LastAssignmentIsUnUsed(allNodes, lastAssignmentNode))
                {
                    unUsedSequentialAssignments.Add(lastAssignmentNode);
                }
            }

            return unUsedSequentialAssignments.Except(usedByNextAssignment);
        }

        /// <summary>
        /// Detects edge case where last AssignmentNode is unused despite the presence of 
        /// a subsequent ReferenceNode. 
        /// </summary>
        /// <example>
        /// <code>
        ///Private Function SpecialCase() As Long
        ///  Dim fizz As Long
        ///  fizz = 0    
        ///     'Edge Case: LHS of fizz = fizz + 1 assignment is unused 
        ///     'since the subsequent ReferenceNode
        ///     'is part on the assignment expression's RHS.  And,
        ///     'the LHS fizz is unused before the end of the procedure
        ///  fizz = fizz + 1
        ///  SpecialCase = 1
        ///End Function
        /// </code>
        /// </example>
        private static bool LastAssignmentIsUnUsed(IEnumerable<INode> allNodes, AssignmentNode lastAssignmentNode)
        {
            var lastAssignmentAndFollowingNodes = allNodes.SkipWhile(n => n != lastAssignmentNode);

            return lastAssignmentAndFollowingNodes.Count() == 2
                        && lastAssignmentAndFollowingNodes.ElementAt(1) is ReferenceNode refNode
                        && AssignmentUsesReferenceNode(lastAssignmentNode, refNode);
        }

        private static IEnumerable<AssignmentNode> DescendantsOfNonEvaluatedTypes(IEnumerable<INode> allNodes)
        {
            var results = new List<AssignmentNode>();
            var allAssignmentNodes = allNodes.Where(n => n.Reference.IsAssignment).Cast<AssignmentNode>();

            foreach (var assignment in allAssignmentNodes)
            {
                if (assignment.TryGetAncestorNode<BranchNode>(out _))
                {
                    results.Add(assignment);
                }
                if (assignment.TryGetAncestorNode<LoopNode>(out _))
                {
                    results.Add(assignment);
                }
            }
            return results;
        }

        private static bool AssignmentUsesReferenceNode(AssignmentNode assignNode, ReferenceNode refNode)
        {
            if (refNode.Reference.Context.TryGetAncestor<VBAParser.LetStmtContext>(out var letAncestor))
            {
                return assignNode.Reference.Context.GetAncestor<VBAParser.LetStmtContext>() == letAncestor;
            }
            if (refNode.Reference.Context.TryGetAncestor<VBAParser.SetStmtContext>(out var setAncestor))
            {
                return assignNode.Reference.Context.GetAncestor<VBAParser.SetStmtContext>() == setAncestor;
            }
            return false;
        }

        //foo = "Hello" <= assignment
        //foo = foo & " World" <= assignment uses prior assignment in RHS expression
        private static bool TryFindAssigmentUsedBySubsequentAssignment(ReferenceNode referenceNode, IEnumerable<AssignmentNode> allAssignments, out AssignmentNode usedAssignment)
        {
            usedAssignment = null;
            //foo = foo & " World"
            var assignmentNode = allAssignments
                                    .FirstOrDefault(an => AssignmentUsesReferenceNode(an, referenceNode));

            if (assignmentNode != null)
            {
                //foo = "Hello"
                usedAssignment = allAssignments.Select(v => v)
                                        .Reverse()
                                        .SkipWhile(n => n != assignmentNode)
                                        .ElementAtOrDefault(1);
            }
            return usedAssignment != null;
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

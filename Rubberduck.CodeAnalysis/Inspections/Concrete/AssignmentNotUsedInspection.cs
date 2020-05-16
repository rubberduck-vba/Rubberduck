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
                        || DisqualifiedByResumeOrGoToStatements(reference, finder));
        }

        private static bool IsAssignmentOfNothing(IdentifierReference reference)
        {
            return reference.IsSetAssignment
                && reference.Context.Parent is VBAParser.SetStmtContext setStmtContext
                && setStmtContext.expression().GetText().Equals(Tokens.Nothing);
        }

        private static bool DisqualifiedByResumeOrGoToStatements(IdentifierReference resultCandidate, DeclarationFinder finder)
        {
            var jumpCtxts = (resultCandidate.ParentScoping.Context.GetDescendents<VBAParser.ResumeStmtContext>().Cast<ParserRuleContext>())
                .Concat(resultCandidate.ParentScoping.Context.GetDescendents<VBAParser.GoToStmtContext>().Cast<ParserRuleContext>());

            if (!jumpCtxts.Any()) { return false; }

            var lineNumbersForNonAssignmentReferences =
                    resultCandidate.Declaration.References
                                                    .Where(rf => !rf.IsAssignment)
                                                    .Select(rf => rf.Context.Stop.Line);

            if (!lineNumbersForNonAssignmentReferences.Any()) { return false; }

            //The jumped-to-line is after the resultCandidate and before a use of the variable 
            return AllJumpToLines(resultCandidate, jumpCtxts, finder)
                                .Any(jumpToLine => resultCandidate.Context.Start.Line > jumpToLine
                                    && lineNumbersForNonAssignmentReferences.Max() >= jumpToLine);
        }

        private static IEnumerable<int> AllJumpToLines(IdentifierReference resultCandidate, IEnumerable<ParserRuleContext> jumpCtxts, DeclarationFinder finder)
        {
            var jumpToLineNumbers = new List<int>();
            var jumpToLabels = new List<string>();
            foreach (var ctxt in jumpCtxts)
            {
                var target = ctxt.children[2].GetText();
                if (int.TryParse(target, out var line))
                {
                    jumpToLineNumbers.Add(line);
                }
                jumpToLabels.Add(target);
            }

            var relevantLabelLineNumbers = finder.DeclarationsWithType(DeclarationType.LineLabel)
                                            .Where(label => resultCandidate.ParentScoping.Equals(label.ParentDeclaration)
                                                                && jumpToLabels.Contains(label.IdentifierName))
                                            .Select(d => d.Context.Start.Line);

            jumpToLineNumbers.AddRange(relevantLabelLineNumbers);

            return jumpToLineNumbers;
        }

        protected override string ResultDescription(IdentifierReference reference)
        {
            return Description;
        }
    }
}

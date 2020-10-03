using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.CodeAnalysis.Inspections.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Extensions;
using Rubberduck.CodeAnalysis.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.Resources.Inspections;
using Rubberduck.VBEditor;
using Rubberduck.CodeAnalysis.Inspections.Concrete.UnreachableCaseEvaluation;

namespace Rubberduck.CodeAnalysis.Inspections.Concrete
{
    /// <summary>
    /// Flags 'Case' blocks that will never execute.
    /// </summary>
    /// <why>
    /// Unreachable code is certainly unintended, and is probably either redundant, or a bug.
    /// </why>
    /// <remarks>
    /// Not all unreachable 'Case' blocks may be flagged, depending on expression complexity.
    /// </remarks>
    /// <example hasresult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Private Sub Example(ByVal value As Long)
    ///     Select Case value
    ///         Case 0 To 99
    ///             ' ...
    ///         Case 50 ' unreachable: case is covered by a preceding condition.
    ///             ' ...
    ///         Case Is < 100
    ///             ' ...
    ///         Case < 0 ' unreachable: case is covered by a preceding condition.
    ///             ' ...
    ///     End Select
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// 
    /// 'If the cumulative result of multiple 'Case' statements
    /// 'cover the entire range of possible values for a data type,
    /// 'then all remaining 'Case' statements are unreachable
    /// 
    /// Private Sub ExampleAllValuesCoveredIntegral(ByVal value As Long, ByVal result As Long)
    ///     Select Case result
    ///         Case Is < 100
    ///             ' ...
    ///         Case Is > -100 
    ///             ' ...
    ///   'all possible values are covered by preceding 'Case' statements 
    ///         Case value * value  ' unreachable
    ///             ' ...
    ///         Case value + value  ' unreachable
    ///             ' ...
    ///         Case Else       ' unreachable 
    ///             ' ...
    ///     End Select
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="false">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// Public Enum ProductID
    ///     Widget = 1
    ///     Gadget = 2
    ///     Gizmo = 3
    /// End Enum
    /// 
    /// Public Sub ExampleEnumCaseElse(ByVal product As ProductID)
    ///
    ///     'Enums are evaluated as the 'Long' data type.  So, in this example,
    ///     'even though all the ProductID enum values have a 'Case' statement, 
    ///     'the 'Case Else' will still execute for any value of the 'product' 
    ///     'parameter that is not a ProductID.
    ///
    ///     Select Case product
    ///         Case Widget
    ///             ' ...
    ///         Case Gadget
    ///             ' ...
    ///         Case Gizmo
    ///             ' ...
    ///         Case Else 'is reachable
    ///             ' Raise an error for unrecognized/unhandled ProductID
    ///     End Select
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    /// <example hasresult="true">
    /// <module name="MyModule" type="Standard Module">
    /// <![CDATA[
    /// 
    /// 'The inspection flags Range Clauses that are not of the required form:
    /// '[x] To [y] where [x] less than or equal to [y]
    /// 
    /// Private Sub ExampleInvalidRangeExpression(ByVal value As String)
    ///     Select Case value
    ///         Case "Beginning" To "End"
    ///             ' ...
    ///         Case "Start" To "Finish" ' unreachable: incorrect form.
    ///             ' ...
    ///         Case Else 
    ///             ' ...
    ///     End Select
    /// End Sub
    /// ]]>
    /// </module>
    /// </example>
    internal sealed class UnreachableCaseInspection : InspectionBase, IParseTreeInspection
    {
        private readonly IUnreachableCaseInspector _inspector;
        private readonly IParseTreeValueVisitor _parseTreeValueVisitor;
        private readonly IInspectionListener<VBAParser.SelectCaseStmtContext> _listener;

        public enum CaseInspectionResultType
        {
            Unreachable,
            InherentlyUnreachable,
            MismatchType,
            Overflow,
            CaseElse
        }

        public UnreachableCaseInspection(
            IDeclarationFinderProvider declarationFinderProvider, 
            IUnreachableCaseInspector inspector, 
            IParseTreeValueVisitor parseTreeValueVisitor) 
            : base(declarationFinderProvider)
        {
            _inspector = inspector;
            _parseTreeValueVisitor = parseTreeValueVisitor;
            _listener = new UnreachableCaseInspectionListener();
        }

        public CodeKind TargetKindOfCode => CodeKind.CodePaneCode;

        public IInspectionListener Listener => _listener;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(DeclarationFinder finder)
        {
            return finder.UserDeclarations(DeclarationType.Module)
                .Where(module => module != null)
                .SelectMany(module => DoGetInspectionResults(module.QualifiedModuleName, finder))
                .ToList();
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var qualifiedSelectCaseStmts = _listener.Contexts(module)
                // ignore filtering here to make the search space smaller
                .Where(result => !result.IsIgnoringInspectionResultFor(finder, AnnotationName));

            return qualifiedSelectCaseStmts
                .SelectMany(context => ResultsForContext(context, finder))
                .ToList();
        }

        private IEnumerable<IInspectionResult> ResultsForContext(QualifiedContext<VBAParser.SelectCaseStmtContext> qualifiedSelectCaseStmt, DeclarationFinder finder)
        {
            var module = qualifiedSelectCaseStmt.ModuleName;
            var selectStmt = qualifiedSelectCaseStmt.Context;
            var contextValues = _parseTreeValueVisitor.VisitChildren(module, selectStmt, finder);

            var results = _inspector.InspectForUnreachableCases(module, selectStmt, contextValues, finder);

            return results
                .Select(resultTpl => CreateInspectionResult(qualifiedSelectCaseStmt, resultTpl.context, resultTpl.resultType))
                .ToList();
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<VBAParser.SelectCaseStmtContext> selectStmt, ParserRuleContext unreachableBlock, CaseInspectionResultType resultType)
        {
            return CreateInspectionResult(selectStmt, unreachableBlock, ResultMessage(resultType));
        }

        //This cannot be a dictionary because the strings have to change after a change in the selected language.
        private static string ResultMessage(CaseInspectionResultType resultType)
        {
            switch (resultType)
            {
                case CaseInspectionResultType.Unreachable:
                    return InspectionResults.UnreachableCaseInspection_Unreachable;
                case CaseInspectionResultType.InherentlyUnreachable:
                    return InspectionResults.UnreachableCaseInspection_InherentlyUnreachable;
                case CaseInspectionResultType.MismatchType:
                    return InspectionResults.UnreachableCaseInspection_TypeMismatch;
                case CaseInspectionResultType.Overflow:
                    return InspectionResults.UnreachableCaseInspection_Overflow;
                case CaseInspectionResultType.CaseElse:
                    return InspectionResults.UnreachableCaseInspection_CaseElse;
                default:
                    throw new ArgumentOutOfRangeException(nameof(resultType), resultType, null);
            }
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<VBAParser.SelectCaseStmtContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            return new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
        }

        private class UnreachableCaseInspectionListener : InspectionListenerBase<VBAParser.SelectCaseStmtContext>
        {
            public override void EnterSelectCaseStmt([NotNull] VBAParser.SelectCaseStmtContext context)
            {
                SaveContext(context);
            }
        }
    }
}
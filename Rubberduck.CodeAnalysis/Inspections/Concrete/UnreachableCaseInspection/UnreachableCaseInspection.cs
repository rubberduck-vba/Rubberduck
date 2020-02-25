using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Resources.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using System;
using Rubberduck.Inspections.Inspections.Extensions;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
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
    /// </example>
    /// <example hasresult="true">
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
    /// </example>
    /// <example hasresult="false">
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
    /// </example>
    /// <example hasresult="true">
    /// <![CDATA[
    /// 
    /// 'The inspecion flags Range Clauses that are not of the required form:
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
    /// </example>
    public sealed class UnreachableCaseInspection : InspectionBase, IParseTreeInspection
    {
        private readonly IUnreachableCaseInspectorFactory _unreachableCaseInspectorFactory;
        private readonly IParseTreeValueFactory _valueFactory;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;

        private enum CaseInspectionResult { Unreachable, InherentlyUnreachable, MismatchType, Overflow, CaseElse };

        private static readonly Dictionary<CaseInspectionResult, string> ResultMessages = new Dictionary<CaseInspectionResult, string>()
        {
            [CaseInspectionResult.Unreachable] = InspectionResults.UnreachableCaseInspection_Unreachable,
            [CaseInspectionResult.InherentlyUnreachable] = InspectionResults.UnreachableCaseInspection_InherentlyUnreachable,
            [CaseInspectionResult.MismatchType] = InspectionResults.UnreachableCaseInspection_TypeMismatch,
            [CaseInspectionResult.Overflow] = InspectionResults.UnreachableCaseInspection_Overflow,
            [CaseInspectionResult.CaseElse] = InspectionResults.UnreachableCaseInspection_CaseElse
        };

        public UnreachableCaseInspection(RubberduckParserState state, IDeclarationFinderProvider declarationFinderProvider) 
            : base(state)
        {
            var factoryProvider = new UnreachableCaseInspectionFactoryProvider();

            _unreachableCaseInspectorFactory = factoryProvider.CreateIUnreachableInspectorFactory();
            _valueFactory = factoryProvider.CreateIParseTreeValueFactory();
            _declarationFinderProvider = declarationFinderProvider;
        }

        public CodeKind TargetKindOfCode => CodeKind.CodePaneCode;

        public IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        private List<IInspectionResult> _inspectionResults = new List<IInspectionResult>();
        private ParseTreeVisitorResults ValueResults { get; }  = new ParseTreeVisitorResults();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            //FIXME Get the declaration finder only once inside the inspection to avoid possible inconsistent state due to a reparse while inspections run.
            _inspectionResults = new List<IInspectionResult>();
            var qualifiedSelectCaseStmts = Listener.Contexts()
                // ignore filtering here to make the search space smaller
                .Where(result => !result.IsIgnoringInspectionResultFor(_declarationFinderProvider.DeclarationFinder, AnnotationName));

            ParseTreeValueVisitor.OnValueResultCreated += ValueResults.OnNewValueResult;

            foreach (var qualifiedSelectCaseStmt in qualifiedSelectCaseStmts)
            {
                qualifiedSelectCaseStmt.Context.Accept(ParseTreeValueVisitor);
                var selectCaseInspector = _unreachableCaseInspectorFactory.Create((VBAParser.SelectCaseStmtContext)qualifiedSelectCaseStmt.Context, ValueResults, _valueFactory, GetVariableTypeName);

                selectCaseInspector.InspectForUnreachableCases();

                selectCaseInspector.UnreachableCases.ForEach(uc => CreateInspectionResult(qualifiedSelectCaseStmt, uc, ResultMessages[CaseInspectionResult.Unreachable]));
                selectCaseInspector.MismatchTypeCases.ForEach(mm => CreateInspectionResult(qualifiedSelectCaseStmt, mm, ResultMessages[CaseInspectionResult.MismatchType]));
                selectCaseInspector.OverflowCases.ForEach(mm => CreateInspectionResult(qualifiedSelectCaseStmt, mm, ResultMessages[CaseInspectionResult.Overflow]));
                selectCaseInspector.InherentlyUnreachableCases.ForEach(mm => CreateInspectionResult(qualifiedSelectCaseStmt, mm, ResultMessages[CaseInspectionResult.InherentlyUnreachable]));
                selectCaseInspector.UnreachableCaseElseCases.ForEach(ce => CreateInspectionResult(qualifiedSelectCaseStmt, ce, ResultMessages[CaseInspectionResult.CaseElse]));
            }
            return _inspectionResults;
        }

        private IParseTreeValueVisitor _parseTreeValueVisitor;
        public IParseTreeValueVisitor ParseTreeValueVisitor
        {
            get
            {
                if (_parseTreeValueVisitor is null)
                {
                    var listener = (UnreachableCaseInspectionListener)Listener;
                    _parseTreeValueVisitor = CreateParseTreeValueVisitor(_valueFactory, listener.EnumerationStmtContexts(), GetIdentifierReferenceForContext);
                }
                return _parseTreeValueVisitor;
            }
        }

        private void CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            var result = new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
            _inspectionResults.Add(result);
        }

        public static IParseTreeValueVisitor CreateParseTreeValueVisitor(IParseTreeValueFactory valueFactory, IReadOnlyList<VBAParser.EnumerationStmtContext> allEnums, Func<ParserRuleContext, (bool success, IdentifierReference idRef)> func)
            => new ParseTreeValueVisitor(valueFactory, allEnums, func);

        //Method is used as a delegate to avoid propagating RubberduckParserState beyond this class
        private (bool success, IdentifierReference idRef) GetIdentifierReferenceForContext(ParserRuleContext context)
        {
            return GetIdentifierReferenceForContext(context, _declarationFinderProvider);
        }

        //public static to support tests
        //FIXME There should not be additional public methods just for tests. This class seems to want to be split or at least reorganized.
        public static (bool success, IdentifierReference idRef) GetIdentifierReferenceForContext(ParserRuleContext context, IDeclarationFinderProvider declarationFinderProvider)
        {
            if (context == null)
            {
                return (false, null);
            }

            var finder = declarationFinderProvider.DeclarationFinder;
            var identifierReferences = finder.MatchName(context.GetText())
                .SelectMany(declaration => declaration.References)
                .Where(reference => reference.Context == context)
                .ToList();

            return identifierReferences.Count == 1 
                ? (true, identifierReferences.First())
                : (false, null);
        }

        //Method is used as a delegate to avoid propogating RubberduckParserState beyond this class
        private string GetVariableTypeName(string variableName, ParserRuleContext ancestor)
        {
            var descendents = ancestor.GetDescendents<VBAParser.SimpleNameExprContext>().Where(desc => desc.GetText().Equals(variableName)).ToList();
            if (descendents.Any())
            {
                (bool success, IdentifierReference idRef) = GetIdentifierReferenceForContext(descendents.First(), _declarationFinderProvider);
                if (success)
                {
                    return GetBaseTypeForDeclaration(idRef.Declaration);
                }
            }
            return string.Empty;
        }

        private string GetBaseTypeForDeclaration(Declaration declaration)
        {
            var localDeclaration = declaration;
            var iterationGuard = 0;
            while (!(localDeclaration is null)
                && !localDeclaration.AsTypeIsBaseType
                && iterationGuard++ < 5)
            {
                localDeclaration = localDeclaration.AsTypeDeclaration;
            }
            return localDeclaration is null ? declaration.AsTypeName : localDeclaration.AsTypeName;
        }

        #region UnreachableCaseInspectionListeners
        public class UnreachableCaseInspectionListener : InspectionListenerBase
        {
            private readonly IDictionary<QualifiedModuleName, List<VBAParser.EnumerationStmtContext>> _enumStmts = new Dictionary<QualifiedModuleName, List<VBAParser.EnumerationStmtContext>>();
            public IReadOnlyList<VBAParser.EnumerationStmtContext> EnumerationStmtContexts() => _enumStmts.AllValues().ToList();
            public IReadOnlyList<VBAParser.EnumerationStmtContext> EnumerationStmtContexts(QualifiedModuleName module) => 
                _enumStmts.TryGetValue(module, out var stmts)
                    ? stmts
                    : new List<VBAParser.EnumerationStmtContext>();

            public override void ClearContexts()
            {
                _enumStmts.Clear();
                base.ClearContexts();
            }

            public override void EnterSelectCaseStmt([NotNull] VBAParser.SelectCaseStmtContext context)
            {
                SaveContext(context);
            }

            public override void EnterEnumerationStmt([NotNull] VBAParser.EnumerationStmtContext context)
            {
                SaveEnumStmt(context);
            }

            private void SaveEnumStmt(VBAParser.EnumerationStmtContext context)
            {
                var module = CurrentModuleName;
                if (_enumStmts.TryGetValue(module, out var stmts))
                {
                    stmts.Add(context);
                }
                else
                {
                    _enumStmts.Add(module, new List<VBAParser.EnumerationStmtContext> { context });
                }
            }
        }
        #endregion
    }
}
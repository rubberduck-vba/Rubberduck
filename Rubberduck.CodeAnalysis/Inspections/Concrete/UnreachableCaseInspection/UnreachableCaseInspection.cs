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
using Rubberduck.JunkDrawer.Extensions;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.Parsing.VBA.Parsing;

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
        private readonly IParseTreeValueVisitorFactory _parseTreeValueVisitorFactory;
        private readonly UnreachableCaseInspectionListener _listener;

        public enum CaseInspectionResultType
        {
            Unreachable,
            InherentlyUnreachable,
            MismatchType,
            Overflow,
            CaseElse
        }

        public UnreachableCaseInspection(IDeclarationFinderProvider declarationFinderProvider, IUnreachableCaseInspectionFactoryProvider factoryProvider) 
            : base(declarationFinderProvider)
        {
            _unreachableCaseInspectorFactory = factoryProvider.CreateIUnreachableInspectorFactory();
            _parseTreeValueVisitorFactory = factoryProvider.CreateParseTreeValueVisitorFactory();
            _listener = new UnreachableCaseInspectionListener();
        }

        public CodeKind TargetKindOfCode => CodeKind.CodePaneCode;

        public IInspectionListener Listener => _listener;

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;

            return finder.UserDeclarations(DeclarationType.Module)
                .Where(module => module != null)
                .SelectMany(module => DoGetInspectionResults(module.QualifiedModuleName, finder))
                .ToList();
        }

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module)
        {
            var finder = DeclarationFinderProvider.DeclarationFinder;
            return DoGetInspectionResults(module, finder);
        }

        private IEnumerable<IInspectionResult> DoGetInspectionResults(QualifiedModuleName module, DeclarationFinder finder)
        {
            var qualifiedSelectCaseStmts = Listener.Contexts(module)
                // ignore filtering here to make the search space smaller
                .Where(result => !result.IsIgnoringInspectionResultFor(finder, AnnotationName));

            var enumStmts = _listener.EnumerationStmtContexts();
            var parseTreeValueVisitor = CreateParseTreeValueVisitor(enumStmts, context => GetIdentifierReferenceForContextFunction(finder)(module, context));

            return qualifiedSelectCaseStmts
                .SelectMany(context => ResultsForContext(context, finder, parseTreeValueVisitor))
                .ToList();
        }

        private IEnumerable<IInspectionResult> ResultsForContext(QualifiedContext<ParserRuleContext> qualifiedSelectCaseStmt, DeclarationFinder finder, IParseTreeValueVisitor parseTreeValueVisitor)
        {
            var module = qualifiedSelectCaseStmt.ModuleName;
            var contextValues = parseTreeValueVisitor.VisitChildren(qualifiedSelectCaseStmt.Context);
            var selectCaseInspector = _unreachableCaseInspectorFactory.Create((VBAParser.SelectCaseStmtContext)qualifiedSelectCaseStmt.Context, contextValues, GetVariableTypeNameFunction(module, finder));

            var results = selectCaseInspector.InspectForUnreachableCases();

            return results
                .Select(resultTpl => CreateInspectionResult(qualifiedSelectCaseStmt, resultTpl.context, resultTpl.resultType))
                .ToList();
        }

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, CaseInspectionResultType resultType)
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

        private IInspectionResult CreateInspectionResult(QualifiedContext<ParserRuleContext> selectStmt, ParserRuleContext unreachableBlock, string message)
        {
            return new QualifiedContextInspectionResult(this,
                        message,
                        new QualifiedContext<ParserRuleContext>(selectStmt.ModuleName, unreachableBlock));
        }

        public IParseTreeValueVisitor CreateParseTreeValueVisitor(
            IReadOnlyList<QualifiedContext<VBAParser.EnumerationStmtContext>> allEnums,
            Func<ParserRuleContext, (bool success, IdentifierReference idRef)> func)
        {
            var enums = allEnums.Select(item => item.Context).ToList();
            return _parseTreeValueVisitorFactory.Create(enums, func);
        }

        private Func<QualifiedModuleName, ParserRuleContext,(bool success, IdentifierReference idRef)> GetIdentifierReferenceForContextFunction(DeclarationFinder finder)
        {
            return (module, context) => GetIdentifierReferenceForContext(module, context, finder);
        }

        //public static to support tests
        //FIXME There should not be additional public methods just for tests. This class seems to want to be split or at least reorganized.
        public static (bool success, IdentifierReference idRef) GetIdentifierReferenceForContext(QualifiedModuleName module, ParserRuleContext context, DeclarationFinder finder)
        {
            if (context == null)
            {
                return (false, null);
            }

            var qualifiedSelection = new QualifiedSelection(module, context.GetSelection());

            var identifierReferences = 
                finder
                .IdentifierReferences(qualifiedSelection)
                .Where(reference => reference.Context == context)
                .ToList();

            return identifierReferences.Count == 1 
                ? (true, identifierReferences.First())
                : (false, null);
        }

        private Func<string, ParserRuleContext, string> GetVariableTypeNameFunction(QualifiedModuleName module, DeclarationFinder finder)
        {
            return (variableName, ancestor) => GetVariableTypeName(module, variableName, ancestor, finder);
        }

        private string GetVariableTypeName(QualifiedModuleName module, string variableName, ParserRuleContext ancestor, DeclarationFinder finder)
        {
            if (ancestor == null)
            {
                return string.Empty;
            }

            var descendents = ancestor.GetDescendents<VBAParser.SimpleNameExprContext>()
                .Where(desc => desc.GetText().Equals(variableName))
                .ToList();
            if (!descendents.Any())
            {
                return string.Empty;
            }

            var firstDescendent = descendents.First();
            var (success, reference) = GetIdentifierReferenceForContext(module, firstDescendent, finder);
            return success ?
                GetBaseTypeForDeclaration(reference.Declaration) 
                : string.Empty;
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
            private readonly IDictionary<QualifiedModuleName, List<QualifiedContext<VBAParser.EnumerationStmtContext>>> _enumStmts = new Dictionary<QualifiedModuleName, List<QualifiedContext<VBAParser.EnumerationStmtContext>>>();
            public IReadOnlyList<QualifiedContext<VBAParser.EnumerationStmtContext>> EnumerationStmtContexts() => _enumStmts.AllValues().ToList();

            public override void ClearContexts()
            {
                _enumStmts.Clear();
                base.ClearContexts();
            }

            public override void ClearContexts(QualifiedModuleName module)
            {
                _enumStmts.Remove(module);
                base.ClearContexts(module);
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
                var qualifiedContext = new QualifiedContext<VBAParser.EnumerationStmtContext>(module, context);
                if (_enumStmts.TryGetValue(module, out var stmts))
                {
                    stmts.Add(qualifiedContext);
                }
                else
                {
                    _enumStmts.Add(module, new List<QualifiedContext<VBAParser.EnumerationStmtContext>> { qualifiedContext });
                }
            }
        }
        #endregion
    }
}
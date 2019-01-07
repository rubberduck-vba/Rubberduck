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

namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public sealed class UnreachableCaseInspection : ParseTreeInspectionBase
    {
        private readonly IUnreachableCaseInspectorFactory _unreachableCaseInspectorFactory;
        private readonly IParseTreeValueFactory _valueFactory;

        private enum CaseInspectionResult { Unreachable, InherentlyUnreachable, MismatchType, Overflow, CaseElse };

        private static readonly Dictionary<CaseInspectionResult, string> ResultMessages = new Dictionary<CaseInspectionResult, string>()
        {
            [CaseInspectionResult.Unreachable] = InspectionResults.UnreachableCaseInspection_Unreachable,
            [CaseInspectionResult.InherentlyUnreachable] = InspectionResults.UnreachableCaseInspection_InherentlyUnreachable,
            [CaseInspectionResult.MismatchType] = InspectionResults.UnreachableCaseInspection_TypeMismatch,
            [CaseInspectionResult.Overflow] = InspectionResults.UnreachableCaseInspection_Overflow,
            [CaseInspectionResult.CaseElse] = InspectionResults.UnreachableCaseInspection_CaseElse
        };

        public UnreachableCaseInspection(RubberduckParserState state) : base(state)
        {
            var factoryProvider = new UnreachableCaseInspectionFactoryProvider();

            _unreachableCaseInspectorFactory = factoryProvider.CreateIUnreachableInspectorFactory();
            _valueFactory = factoryProvider.CreateIParseTreeValueFactory();
        }

        public override IInspectionListener Listener { get; } =
            new UnreachableCaseInspectionListener();

        private List<IInspectionResult> _inspectionResults = new List<IInspectionResult>();
        private ParseTreeVisitorResults ValueResults { get; }  = new ParseTreeVisitorResults();

        protected override IEnumerable<IInspectionResult> DoGetInspectionResults()
        {
            _inspectionResults = new List<IInspectionResult>();
            var qualifiedSelectCaseStmts = Listener.Contexts
                .Where(result => !IsIgnoringInspectionResultFor(result.ModuleName, result.Context.Start.Line));

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
                    _parseTreeValueVisitor = CreateParseTreeValueVisitor(_valueFactory, listener.EnumerationStmtContexts.ToList(), GetIdentifierReferenceForContext);
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

        public static IParseTreeValueVisitor CreateParseTreeValueVisitor(IParseTreeValueFactory valueFactory, List<VBAParser.EnumerationStmtContext> allEnums, Func<ParserRuleContext, (bool success, IdentifierReference idRef)> func)
            => new ParseTreeValueVisitor(valueFactory, allEnums, func);

        //Method is used as a delegate to avoid propogating RubberduckParserState beyond this class
        private (bool success, IdentifierReference idRef) GetIdentifierReferenceForContext(ParserRuleContext context)
        {
            return GetIdentifierReferenceForContext(context, State);
        }

        //public static to support tests
        public static (bool success, IdentifierReference idRef) GetIdentifierReferenceForContext(ParserRuleContext context, RubberduckParserState state)
        {
            IdentifierReference idRef = null;
            var success = false;
            var identifierReferences = (state.DeclarationFinder.MatchName(context.GetText()).Select(dec => dec.References)).SelectMany(rf => rf)
                .Where(rf => rf.Context == context);
            if (identifierReferences.Count() == 1)
            {
                idRef = identifierReferences.First();
                success = true;
            }
            return (success, idRef);
        }

        //Method is used as a delegate to avoid propogating RubberduckParserState beyond this class
        private string GetVariableTypeName(string variableName, ParserRuleContext ancestor)
        {
            var descendents = ancestor.GetDescendents<VBAParser.SimpleNameExprContext>().Where(desc => desc.GetText().Equals(variableName));
            if (descendents.Any())
            {
                (bool success, IdentifierReference idRef) = GetIdentifierReferenceForContext(descendents.First(), State);
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
        public class UnreachableCaseInspectionListener : VBAParserBaseListener, IInspectionListener
        {
            private readonly List<QualifiedContext<ParserRuleContext>> _contexts = new List<QualifiedContext<ParserRuleContext>>();
            public IReadOnlyList<QualifiedContext<ParserRuleContext>> Contexts => _contexts;

            private readonly List<VBAParser.EnumerationStmtContext> _enumStmts = new List<VBAParser.EnumerationStmtContext>();
            public IReadOnlyList<VBAParser.EnumerationStmtContext> EnumerationStmtContexts => _enumStmts;

            public QualifiedModuleName CurrentModuleName { get; set; }

            public void ClearContexts()
            {
                _contexts.Clear();
            }

            public override void EnterSelectCaseStmt([NotNull] VBAParser.SelectCaseStmtContext context)
            {
                _contexts.Add(new QualifiedContext<ParserRuleContext>(CurrentModuleName, context));
            }

            public override void EnterEnumerationStmt([NotNull] VBAParser.EnumerationStmtContext context)
            {
                _enumStmts.Add(context);
            }
        }
        #endregion
    }
}
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Antlr4.Runtime;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA.DeclarationCaching;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.VBA.ReferenceManagement
{
    public sealed class IdentifierReferenceResolver
    {
        private readonly DeclarationFinder _declarationFinder;
        private readonly QualifiedModuleName _qualifiedModuleName;
        private readonly Stack<IBoundExpression> _withBlockExpressions;
        private readonly Declaration _moduleDeclaration;
        private Declaration _currentScope;
        private Declaration _currentParent;
        private readonly BindingService _bindingService;
        private readonly BoundExpressionVisitor _boundExpressionVisitor;
        private readonly FailedResolutionVisitor _failedResolutionVisitor;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public IdentifierReferenceResolver(QualifiedModuleName qualifiedModuleName, DeclarationFinder finder)
        {
            _declarationFinder = finder;
            _qualifiedModuleName = qualifiedModuleName;
            _withBlockExpressions = new Stack<IBoundExpression>();
            _moduleDeclaration = finder.MatchName(_qualifiedModuleName.ComponentName)
                .SingleOrDefault(item =>
                    (item.DeclarationType.HasFlag(DeclarationType.ClassModule) ||
                     item.DeclarationType == DeclarationType.ProceduralModule)
                    && item.QualifiedName.QualifiedModuleName.Equals(_qualifiedModuleName));
            SetCurrentScope();
            var typeBindingContext = new TypeBindingContext(_declarationFinder);
            var procedurePointerBindingContext = new ProcedurePointerBindingContext(_declarationFinder);
            _bindingService = new BindingService(
                _declarationFinder,
                new DefaultBindingContext(_declarationFinder, typeBindingContext, procedurePointerBindingContext),
                typeBindingContext,
                procedurePointerBindingContext);
            _boundExpressionVisitor = new BoundExpressionVisitor(finder);
            _failedResolutionVisitor = new FailedResolutionVisitor(finder);
        }

        public void SetCurrentScope()
        {
            _currentScope = _moduleDeclaration;
            _currentParent = _moduleDeclaration;
        }

        public void SetCurrentScope(string memberName, DeclarationType type)
        {
            Logger.Trace("Setting current scope: {0} ({1}) in thread {2}", memberName, type,
                Thread.CurrentThread.ManagedThreadId);

            _currentParent = _declarationFinder.MatchName(memberName).SingleOrDefault(item =>
                item.QualifiedName.QualifiedModuleName == _qualifiedModuleName && item.DeclarationType == type);

            _currentScope = _declarationFinder.MatchName(memberName).SingleOrDefault(item =>
                item.QualifiedName.QualifiedModuleName == _qualifiedModuleName && item.DeclarationType == type) ??
                            _moduleDeclaration;

            Logger.Trace("Current scope is now {0} in thread {1}",
                _currentScope == null ? "null" : _currentScope.IdentifierName, Thread.CurrentThread.ManagedThreadId);
        }

        public void EnterWithBlock(VBAParser.WithStmtContext context)
        {
            var withExpression = GetInnerMostWithExpression();
            var boundExpression = _bindingService.ResolveDefault(
                _moduleDeclaration,
                _currentParent,
                context.expression(),
                withExpression,
                StatementResolutionContext.Undefined,
                false);
            _failedResolutionVisitor.CollectUnresolved(boundExpression, _currentParent, withExpression);
            _boundExpressionVisitor.AddIdentifierReferences(boundExpression, _qualifiedModuleName, _currentScope, _currentParent);
            _withBlockExpressions.Push(boundExpression);
        }

        private IBoundExpression GetInnerMostWithExpression()
        {
            if (_withBlockExpressions.Any())
            {
                return _withBlockExpressions.Peek();
            }
            return null;
        }

        public void ExitWithBlock()
        {
            _withBlockExpressions.Pop();
        }

        public void Resolve(VBAParser.ArgDefaultValueContext context)
        {
            var expression = context.expression();
            if (expression == null)
            {
                return;
            }
            ResolveDefault(expression, false);
        }

        public void Resolve(VBAParser.ArrayDimContext context)
        {
            if (context.boundsList() == null)
            {
                return;
            }
            foreach (var dimSpec in context.boundsList().dimSpec())
            {
                if (dimSpec.lowerBound() != null)
                {
                    ResolveDefault(dimSpec.lowerBound().constantExpression().expression(), true);
                }
                ResolveDefault(dimSpec.upperBound().constantExpression().expression(), true);
            }
        }

        public void Resolve(VBAParser.OnErrorStmtContext context)
        {
            if (context.expression() == null)
            {
                return;
            }
            ResolveLabel(context.expression(), context.expression().GetText());
        }

        public void Resolve(VBAParser.ErrorStmtContext context)
        {
            ResolveDefault(context.expression(), true);
        }

        private void ResolveLabel(ParserRuleContext context, string label)
        {
            var labelDeclaration = _bindingService.ResolveGoTo(_currentParent, label);
            if (labelDeclaration != null)
            {
                var callSiteContext = context;
                var identifier = context.GetText();
                var callee = labelDeclaration;
                labelDeclaration.AddReference(
                    _qualifiedModuleName,
                    _currentScope,
                    _currentParent,
                    callSiteContext,
                    identifier,
                    callee,
                    callSiteContext.GetSelection(),
                    FindIdentifierAnnotations(_qualifiedModuleName, callSiteContext.GetSelection().StartLine));
            }
        }

        private IEnumerable<IParseTreeAnnotation> FindIdentifierAnnotations(QualifiedModuleName module, int line)
        {
            return _declarationFinder.FindAnnotations(module, line, AnnotationTarget.Identifier);
        }

        private void ResolveDefault(
            ParserRuleContext expression,
            bool requiresLetCoercion = false,
            StatementResolutionContext statementContext = StatementResolutionContext.Undefined,
            bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false,
            bool isSetAssignment = false,
            bool isReDim = false)
        {
            var withExpression = GetInnerMostWithExpression();
            var boundExpression = _bindingService.ResolveDefault(
                _moduleDeclaration,
                _currentParent,
                expression,
                withExpression,
                statementContext,
                requiresLetCoercion,
                isAssignmentTarget);

            _failedResolutionVisitor.CollectUnresolved(boundExpression, _currentParent, withExpression);

            _boundExpressionVisitor.AddIdentifierReferences(
                boundExpression, 
                _qualifiedModuleName, 
                _currentScope,
                _currentParent,
                isAssignmentTarget,
                hasExplicitLetStatement, 
                isSetAssignment,
                isReDim);
        }

        private void ResolveType(ParserRuleContext expression)
        {
            var boundExpression = _bindingService.ResolveType(_moduleDeclaration, _currentParent, expression);
            if (boundExpression.Classification == ExpressionClassification.ResolutionFailed)
            {
                Logger.Warn($"Type Context: Failed to resolve {expression.GetText()}. Binding as much as we can.");
            }

            _failedResolutionVisitor.CollectUnresolved(boundExpression, _currentParent, GetInnerMostWithExpression());
            _boundExpressionVisitor.AddIdentifierReferences(boundExpression, _qualifiedModuleName, _currentScope, _currentParent);
        }

        public void Resolve(VBAParser.GoToStmtContext context)
        {
            ResolveLabel(context.expression(), context.expression().GetText());
        }

        public void Resolve(VBAParser.OnGoToStmtContext context)
        {
            ResolveDefault(context.expression()[0], true);
            for (int labelIndex = 1; labelIndex < context.expression().Length; labelIndex++)
            {
                ResolveLabel(context.expression()[labelIndex], context.expression()[labelIndex].GetText());
            }
        }

        public void Resolve(VBAParser.GoSubStmtContext context)
        {
            ResolveLabel(context.expression(), context.expression().GetText());
        }

        public void Resolve(VBAParser.OnGoSubStmtContext context)
        {
            ResolveDefault(context.expression()[0], true);
            for (int labelIndex = 1; labelIndex < context.expression().Length; labelIndex++)
            {
                ResolveLabel(context.expression()[labelIndex], context.expression()[labelIndex].GetText());
            }
        }

        public void Resolve(VBAParser.RedimStmtContext context)
        {
            // TODO: Create local variable if no match for ReDim variable declaration.
            foreach (var redimVariableDeclaration in context.redimDeclarationList().redimVariableDeclaration())
            {
                // We treat ReDim statements as index expressions to make it SLL.
                var lExpr = ((VBAParser.LExprContext)redimVariableDeclaration.expression()).lExpression();

                VBAParser.LExpressionContext indexedExpression;
                VBAParser.ArgumentListContext argumentList;
                if (lExpr is VBAParser.IndexExprContext indexExpr)
                {
                    indexedExpression = indexExpr.lExpression();
                    argumentList = indexExpr.argumentList();
                }
                else
                {
                    var whitespaceIndexExpr = (VBAParser.WhitespaceIndexExprContext) lExpr;
                    indexedExpression = whitespaceIndexExpr.lExpression();
                    argumentList = whitespaceIndexExpr.argumentList();

                }
                // The indexedExpression is the array that is being resized.
                // We can't treat it as a normal index expression because the semantics are different.
                // It's not actually a function call but a special statement.
                ResolveDefault(indexedExpression, false, isReDim: true);
                if (argumentList.argument() != null)
                {
                    foreach (var positionalArgument in argumentList.argument())
                    {
                        if (positionalArgument.positionalArgument() != null)
                        {
                            ResolveReDimArgument(positionalArgument.positionalArgument().argumentExpression());
                        }
                    }
                }
            }
        }

        private void ResolveReDimArgument(VBAParser.ArgumentExpressionContext argument)
        {
            // ReDim statements can either have "normal" positional argument expressions or lower + upper bounds arguments.
            if (argument.lowerBoundArgumentExpression() != null)
            {
                ResolveDefault(argument.lowerBoundArgumentExpression().expression(), true);
                ResolveDefault(argument.upperBoundArgumentExpression().expression(), true);
            }
            else
            {
                ResolveDefault(argument.expression(), true);
            }
        }

        public void Resolve(VBAParser.WhileWendStmtContext context)
        {
            ResolveDefault(context.expression(), true);
        }

        public void Resolve(VBAParser.DoLoopStmtContext context)
        {
            if (context.expression() == null)
            {
                return;
            }
            ResolveDefault(context.expression(), true);
        }

        public void Resolve(VBAParser.IfStmtContext context)
        {
            ResolveDefault(context.booleanExpression());
            if (context.elseIfBlock() != null)
            {
                foreach (var elseIfBlock in context.elseIfBlock())
                {
                    ResolveDefault(elseIfBlock.booleanExpression());
                }
            }
        }

        public void Resolve(VBAParser.SingleLineIfStmtContext context)
        {
            // The listOrLabel rule could be resolved separately but since it's such a special case, only appearing in
            // single-line-if-statements, we do it here for better understanding.
            if (context.ifWithEmptyThen() != null)
            {
                ResolveDefault(context.ifWithEmptyThen().booleanExpression());
                ResolveListOrLabel(context.ifWithEmptyThen().singleLineElseClause().listOrLabel());
            }
            else
            {
                ResolveDefault(context.ifWithNonEmptyThen().booleanExpression());
                ResolveListOrLabel(context.ifWithNonEmptyThen().listOrLabel());
                if (context.ifWithNonEmptyThen().singleLineElseClause() != null)
                {
                    ResolveListOrLabel(context.ifWithNonEmptyThen().singleLineElseClause().listOrLabel());
                }
            }
        }

        private void ResolveListOrLabel(VBAParser.ListOrLabelContext listOrLabel)
        {
            if (listOrLabel?.lineNumberLabel() == null)
            {
                return;
            }
            ResolveLabel(listOrLabel.lineNumberLabel(), listOrLabel.lineNumberLabel().GetText());
        }

        public void Resolve(VBAParser.SelectCaseStmtContext context)
        {
            ResolveDefault(context.selectExpression().expression(), true);
            if (context.caseClause() == null)
            {
                return;
            }
            foreach (var caseClause in context.caseClause())
            {
                foreach (var rangeClause in caseClause.rangeClause())
                {
                    if (rangeClause.expression() != null)
                    {
                        ResolveDefault(rangeClause.expression(), true);
                    }
                    else
                    {
                        ResolveDefault(rangeClause.selectStartValue().expression(), true);
                        ResolveDefault(rangeClause.selectEndValue().expression(), true);
                    }
                }
            }
        }

        public void Resolve(VBAParser.LetStmtContext context)
        {
            var letStatement = context.LET();
            ResolveDefault(
                context.lExpression(),
                true,
                StatementResolutionContext.LetStatement,
                true,
                letStatement != null);
            ResolveDefault(context.expression(), true);
        }

        public void Resolve(VBAParser.SetStmtContext context)
        {
            ResolveDefault(
                context.lExpression(),
                false,
                StatementResolutionContext.SetStatement,
                true,
                false,
                true);
            ResolveDefault(context.expression(), false);
        }

        public void Resolve(VBAParser.CallStmtContext context)
        {
            ResolveDefault(context, false);
        }

        public void Resolve(VBAParser.ConstStmtContext context)
        {
            foreach (var constStmt in context.constSubStmt())
            {
                ResolveDefault(constStmt.expression(), false);
            }
        }

        public void Resolve(VBAParser.EraseStmtContext context)
        {
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr, true);
            }
        }

        public void Resolve(VBAParser.NameStmtContext context)
        {
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr, true);
            }
        }

        private void ResolveFileNumber(VBAParser.FileNumberContext fileNumber)
        {
            ResolveDefault(fileNumber.markedFileNumber() != null
                ? fileNumber.markedFileNumber().expression()
                : fileNumber.unmarkedFileNumber().expression(),
                true);
        }

        public void Resolve(VBAParser.OpenStmtContext context)
        {
            ResolveDefault(context.pathName().expression(), true);
            ResolveFileNumber(context.fileNumber());
            if (context.lenClause() != null)
            {
                ResolveDefault(context.lenClause().recLength().expression(), true);
            }
        }

        public void Resolve(VBAParser.CloseStmtContext context)
        {
            if (context.fileNumberList() != null)
            {
                foreach (var fileNumber in context.fileNumberList().fileNumber())
                {
                    ResolveFileNumber(fileNumber);
                }
            }
        }

        public void Resolve(VBAParser.SeekStmtContext context)
        {
            ResolveFileNumber(context.fileNumber());
            ResolveDefault(context.position().expression(), true);
        }

        public void Resolve(VBAParser.LockStmtContext context)
        {
            ResolveFileNumber(context.fileNumber());
            if (context.recordRange() == null)
            {
                return;
            }
            ResolveRecordRange(context.recordRange());
        }

        public void Resolve(VBAParser.UnlockStmtContext context)
        {
            ResolveFileNumber(context.fileNumber());
            ResolveRecordRange(context.recordRange());
        }

        private void ResolveRecordRange(VBAParser.RecordRangeContext recordRange)
        {
            if (recordRange == null)
            {
                return;
            }
            if (recordRange.startRecordNumber() != null)
            {
                ResolveDefault(recordRange.startRecordNumber().expression(), true);
            }
            if (recordRange.endRecordNumber() != null)
            {
                ResolveDefault(recordRange.endRecordNumber().expression(), true);
            }
        }

        public void Resolve(VBAParser.LineInputStmtContext context)
        {
            ResolveDefault(context.markedFileNumber().expression(), true);
            ResolveDefault(context.variableName().expression(),false , isAssignmentTarget: true);
        }

        public void Resolve(VBAParser.WidthStmtContext context)
        {
            ResolveDefault(context.markedFileNumber().expression(), true);
            ResolveDefault(context.lineWidth().expression(), true);
        }

        public void Resolve(VBAParser.PrintStmtContext context)
        {
            ResolveDefault(context.markedFileNumber().expression(), true);
            var outputList = context.outputList();
            if (outputList != null)
            {
                ResolveOutputList(outputList);
            }
        }

        public void Resolve(VBAParser.UnqualifiedObjectPrintStmtContext context)
        {
            ResolveDefault(context);
        }

        public void Resolve(VBAParser.WriteStmtContext context)
        {
            ResolveDefault(context.markedFileNumber().expression(), true);
            var outputList = context.outputList();
            if (outputList != null)
            {
                ResolveOutputList(outputList);
            }
        }

        private void ResolveOutputList(VBAParser.OutputListContext outputList)
        {
            ResolveDefault(outputList);
        }

        public void Resolve(VBAParser.InputStmtContext context)
        {
            ResolveDefault(context.markedFileNumber().expression(), true);
            foreach (var inputVariable in context.inputList().inputVariable())
            {
                ResolveDefault(inputVariable.expression(), false, isAssignmentTarget: true);
            }
        }

        public void Resolve(VBAParser.PutStmtContext context)
        {
            ResolveFileNumber(context.fileNumber());
            if (context.recordNumber() != null)
            {
                ResolveDefault(context.recordNumber().expression(), true);
            }
            if (context.data() != null)
            {
                ResolveDefault(context.data().expression(), false);
            }
        }

        public void Resolve(VBAParser.GetStmtContext context)
        {
            ResolveFileNumber(context.fileNumber());
            if (context.recordNumber() != null)
            {
                ResolveDefault(context.recordNumber().expression(), true);
            }
            if (context.variable() != null)
            {
                ResolveDefault(context.variable().expression(), false, isAssignmentTarget: true);
            }
        }

        public void Resolve(VBAParser.LsetStmtContext context)
        {
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr, true);
            }
        }

        public void Resolve(VBAParser.RsetStmtContext context)
        {
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr, true);
            }
        }

        public void Resolve(VBAParser.MidStatementContext context)
        {
            var variableExpression = context.lExpression();
            ResolveDefault(variableExpression, true);
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr, true);
            }
        }

        public void Resolve(VBAParser.AsTypeClauseContext context)
        {
            // All "As Type" expressions are resolved here, statements don't have to resolve their "As Types" themselves.
            var asType = context.type();
            if (asType == null)
            {
                return;
            }
            var baseType = asType.baseType();
            if (baseType != null)
            {
                // Fixed-Length strings can have a constant-name as length that is a simple-name-expression that also has to be resolved.
                var length = context.fieldLength();
                if (length?.identifierValue() != null)
                {
                    ResolveDefault(length.identifierValue(), false);
                }
                return;
            }
            ResolveType(asType.complexType());
        }

        public void Resolve(VBAParser.ForNextStmtContext context)
        {
            var expressions = context.expression();

            // In "For expr1 = expr2" the "expr1 = expr2" part is treated as a single expression.
            var assignmentExpr = ((VBAParser.RelationalOpContext)expressions[0]);
            ResolveStartValueAssignmentOfForNext(assignmentExpr);

            ResolveToValueOfForNext(expressions[1]);

            var stepStatement = context.stepStmt();
            if (stepStatement != null)
            {
                Resolve(stepStatement);
            }

            const int firstNextExpressionIndex = 2;
            for (var exprIndex = firstNextExpressionIndex; exprIndex < expressions.Length; exprIndex++)
            {
                ResolveDefault(expressions[exprIndex]);
            }
        }

        private void ResolveStartValueAssignmentOfForNext(VBAParser.RelationalOpContext expression)
        {
            var expressions = expression.expression();
            var elementVariableExpression = expressions[0];
            ResolveDefault(elementVariableExpression, requiresLetCoercion: true, isAssignmentTarget: true);

            var startValueExpression = expressions[1];
            ResolveDefault(startValueExpression, requiresLetCoercion: true);
        }

        private void ResolveToValueOfForNext(ParserRuleContext expression)
        {
            ResolveDefault(expression, requiresLetCoercion: true);
        }

        private void Resolve(VBAParser.StepStmtContext context)
        {
            ResolveDefault(context.expression(), true);
        }

        public void Resolve(VBAParser.ForEachStmtContext context)
        {
            var expressions = context.expression();

            var elementVariableExpression = expressions[0];
            ResolveDefault(elementVariableExpression, isAssignmentTarget: true);

            var collectionExpression = expressions[1];
            ResolveDefault(collectionExpression);

            const int firstNextExpressionIndex = 2;
            for (var exprIndex = firstNextExpressionIndex; exprIndex < context.expression().Length; exprIndex++)
            {
                ResolveDefault(expressions[exprIndex]);
            }
        }

        public void Resolve(VBAParser.ImplementsStmtContext context)
        {
            ResolveType(context.expression());
        }

        public void Resolve(VBAParser.RaiseEventStmtContext context)
        {
            var eventDeclaration = _bindingService.ResolveEvent(_moduleDeclaration, context.identifier().GetText());
            if (eventDeclaration != null)
            {
                var callSiteContext = context.identifier();
                var identifier = Identifier.GetName(context.identifier());
                var callee = eventDeclaration;
                callee.AddReference(
                    _qualifiedModuleName,
                    _currentScope,
                    _currentParent,
                    callSiteContext,
                    identifier,
                    callee,
                    callSiteContext.GetSelection(),
                    FindIdentifierAnnotations(_qualifiedModuleName, callSiteContext.GetSelection().StartLine));
            }
            if (context.eventArgumentList() == null)
            {
                return;
            }
            foreach (var eventArgument in context.eventArgumentList().eventArgument())
            {
                ResolveDefault(eventArgument.expression(), false);
            }
        }

        public void Resolve(VBAParser.ResumeStmtContext context)
        {
            if (context.expression() == null)
            {
                return;
            }
            ResolveLabel(context.expression(), context.expression().GetText());
        }

        public void Resolve(VBAParser.LineSpecialFormContext context)
        {
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr, true);
            }
            foreach (var tuple in context.tuple())
            {
                ResolveTuple(tuple);
            }
        }

        public void Resolve(VBAParser.CircleSpecialFormContext context)
        {
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr, true);
            }
            ResolveTuple(context.tuple());
        }

        public void Resolve(VBAParser.ScaleSpecialFormContext context)
        {
            if (context.expression() != null)
            {
                ResolveDefault(context.expression(), true);

            }
            foreach (var tuple in context.tuple())
            {
                ResolveTuple(tuple);
            }
        }

        public void Resolve(VBAParser.PSetSpecialFormContext context)
        {
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr, true);
            }
            ResolveTuple(context.tuple());
        }

        private void ResolveTuple(VBAParser.TupleContext tuple)
        {
            foreach (var expr in tuple.expression())
            {
                ResolveDefault(expr, true);
            }
        }

        public void Resolve(VBAParser.EnumerationStmtContext context)
        {
            if (context.enumerationStmt_Constant() == null)
            {
                return;
            }
            foreach (var enumMember in context.enumerationStmt_Constant())
            {
                if (enumMember.expression() != null)
                {
                    ResolveDefault(enumMember.expression(), false);
                }
            }
        }
    }
}

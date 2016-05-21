using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading;

namespace Rubberduck.Parsing.Symbols
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
        private readonly AnnotationService _annotationService;

        public IdentifierReferenceResolver(QualifiedModuleName qualifiedModuleName, DeclarationFinder finder)
        {
            _declarationFinder = finder;
            _qualifiedModuleName = qualifiedModuleName;
            _withBlockExpressions = new Stack<IBoundExpression>();
            _moduleDeclaration = finder.MatchName(_qualifiedModuleName.ComponentName)
                .SingleOrDefault(item =>
                    (item.DeclarationType == DeclarationType.ClassModule ||
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
            _annotationService = new AnnotationService(_declarationFinder);
            _boundExpressionVisitor = new BoundExpressionVisitor(_annotationService);
        }

        public void SetCurrentScope()
        {
            _currentScope = _moduleDeclaration;
            _currentParent = _moduleDeclaration;
        }

        public void SetCurrentScope(string memberName, DeclarationType type)
        {
            Debug.WriteLine("Setting current scope: {0} ({1}) in thread {2}", memberName, type,
                Thread.CurrentThread.ManagedThreadId);

            _currentParent = _declarationFinder.MatchName(memberName).SingleOrDefault(item =>
                item.QualifiedName.QualifiedModuleName == _qualifiedModuleName && item.DeclarationType == type);

            _currentScope = _declarationFinder.MatchName(memberName).SingleOrDefault(item =>
                item.QualifiedName.QualifiedModuleName == _qualifiedModuleName && item.DeclarationType == type) ??
                            _moduleDeclaration;

            Debug.WriteLine("Current scope is now {0} in thread {1}",
                _currentScope == null ? "null" : _currentScope.IdentifierName, Thread.CurrentThread.ManagedThreadId);
        }

        public void EnterWithBlock(VBAParser.WithStmtContext context)
        {
            var boundExpression = _bindingService.ResolveDefault(
                _moduleDeclaration,
                _currentParent,
                context.expression(),
                GetInnerMostWithExpression(),
                StatementResolutionContext.Undefined);
            _boundExpressionVisitor.AddIdentifierReferences(boundExpression, _qualifiedModuleName, _currentScope, _currentParent);
            // note: pushes null if unresolved
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
            ResolveDefault(context.expression());
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
                    _annotationService.FindAnnotations(_qualifiedModuleName, callSiteContext.GetSelection().StartLine));
            }
        }

        private void ResolveDefault(
            ParserRuleContext expression,
            StatementResolutionContext statementContext = StatementResolutionContext.Undefined,
            bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false)
        {
            var boundExpression = _bindingService.ResolveDefault(
                _moduleDeclaration,
                _currentParent,
                expression,
                GetInnerMostWithExpression(),
                statementContext);
            if (boundExpression.Classification == ExpressionClassification.ResolutionFailed)
            {
                Debug.WriteLine(
                    string.Format(
                        "{0}/Default Context: Failed to resolve {1}. Binding all successfully resolved expressions anyway.",
                        GetType().Name,
                        expression.GetText()));
            }
            _boundExpressionVisitor.AddIdentifierReferences(boundExpression, _qualifiedModuleName, _currentScope, _currentParent, isAssignmentTarget, false);
        }

        private void ResolveType(ParserRuleContext expression)
        {
            var boundExpression = _bindingService.ResolveType(_moduleDeclaration, _currentParent, expression);
            if (boundExpression.Classification == ExpressionClassification.ResolutionFailed)
            {
                Debug.WriteLine(
                    string.Format(
                        "{0}/Type Context: Failed to resolve {1}. Binding all successfully resolved expressions anyway.",
                        GetType().Name,
                        expression.GetText()));
            }
            _boundExpressionVisitor.AddIdentifierReferences(boundExpression, _qualifiedModuleName, _currentScope, _currentParent);
        }

        public void Resolve(VBAParser.GoToStmtContext context)
        {
            ResolveLabel(context.expression(), context.expression().GetText());
        }

        public void Resolve(VBAParser.OnGoToStmtContext context)
        {
            ResolveDefault(context.expression()[0]);
            for (int labelIndex = 1; labelIndex < context.expression().Count; labelIndex++)
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
            ResolveDefault(context.expression()[0]);
            for (int labelIndex = 1; labelIndex < context.expression().Count; labelIndex++)
            {
                ResolveLabel(context.expression()[labelIndex], context.expression()[labelIndex].GetText());
            }
        }

        public void Resolve(VBAParser.RedimStmtContext context)
        {
            foreach (var redimVariableExpression in context.redimDeclarationList().redimVariableDeclarationExpression())
            {
                ResolveDefault(redimVariableExpression.lExpression());
                foreach (var dimSpec in redimVariableExpression.dynamicArrayClause().dynamicArrayDim().dynamicBoundsList().dynamicDimSpec())
                {
                    if (dimSpec.dynamicLowerBound() != null)
                    {
                        ResolveDefault(dimSpec.dynamicLowerBound().integerExpression());
                    }
                    ResolveDefault(dimSpec.dynamicUpperBound().integerExpression());
                }
            }
        }

        public void Resolve(VBAParser.WhileWendStmtContext context)
        {
            ResolveDefault(context.expression());
        }

        public void Resolve(VBAParser.DoLoopStmtContext context)
        {
            if (context.expression() == null)
            {
                return;
            }
            ResolveDefault(context.expression());
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
            if (listOrLabel == null || listOrLabel.lineNumberLabel() == null)
            {
                return;
            }
            ResolveLabel(listOrLabel.lineNumberLabel(), listOrLabel.lineNumberLabel().GetText());
        }

        public void Resolve(VBAParser.SelectCaseStmtContext context)
        {
            ResolveDefault(context.selectExpression().expression());
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
                        ResolveDefault(rangeClause.expression());
                    }
                    else
                    {
                        ResolveDefault(rangeClause.selectStartValue().expression());
                        ResolveDefault(rangeClause.selectEndValue().expression());
                    }
                }
            }
        }

        public void Resolve(VBAParser.LetStmtContext context)
        {
            var letStatement = context.LET();
            ResolveDefault(
                context.lExpression(),
                StatementResolutionContext.LetStatement,
                true,
                letStatement != null);
            ResolveDefault(context.expression());
        }

        public void Resolve(VBAParser.SetStmtContext context)
        {
            ResolveDefault(
                context.lExpression(),
                StatementResolutionContext.SetStatement,
                true,
                false);
            ResolveDefault(context.expression());
        }

        public void Resolve(VBAParser.CallStmtContext context)
        {
            ResolveDefault(context);
        }

        public void Resolve(VBAParser.ConstStmtContext context)
        {
            foreach (var constStmt in context.constSubStmt())
            {
                ResolveDefault(constStmt.expression());
            }
        }

        public void Resolve(VBAParser.EraseStmtContext context)
        {
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr);
            }
        }

        private void ResolveFileNumber(VBAParser.FileNumberContext fileNumber)
        {
            if (fileNumber.markedFileNumber() != null)
            {
                ResolveDefault(fileNumber.markedFileNumber().expression());
            }
            else
            {
                ResolveDefault(fileNumber.unmarkedFileNumber().expression());
            }
        }

        public void Resolve(VBAParser.OpenStmtContext context)
        {
            ResolveDefault(context.pathName().expression());
            ResolveFileNumber(context.fileNumber());
            if (context.lenClause() != null)
            {
                ResolveDefault(context.lenClause().recLength().expression());
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
            ResolveDefault(context.position().expression());
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
                ResolveDefault(recordRange.startRecordNumber().expression());
            }
            if (recordRange.endRecordNumber() != null)
            {
                ResolveDefault(recordRange.endRecordNumber().expression());
            }
        }

        public void Resolve(VBAParser.LineInputStmtContext context)
        {
            ResolveDefault(context.markedFileNumber().expression());
            ResolveDefault(context.variableName().expression());
        }

        public void Resolve(VBAParser.WidthStmtContext context)
        {
            ResolveDefault(context.markedFileNumber().expression());
            ResolveDefault(context.lineWidth().expression());
        }

        public void Resolve(VBAParser.PrintStmtContext context)
        {
            ResolveDefault(context.markedFileNumber().expression());
            ResolveOutputList(context.outputList());
        }

        public void Resolve(VBAParser.WriteStmtContext context)
        {
            ResolveDefault(context.markedFileNumber().expression());
            ResolveOutputList(context.outputList());
        }

        private void ResolveOutputList(VBAParser.OutputListContext outputList)
        {
            if (outputList == null)
            {
                return;
            }
            foreach (var outputItem in outputList.outputItem())
            {
                if (outputItem.outputClause() != null)
                {
                    if (outputItem.outputClause().spcClause() != null)
                    {
                        ResolveDefault(outputItem.outputClause().spcClause().spcNumber().expression());
                    }
                    if (outputItem.outputClause().tabClause() != null && outputItem.outputClause().tabClause().tabNumberClause() != null)
                    {
                        ResolveDefault(outputItem.outputClause().tabClause().tabNumberClause().tabNumber().expression());
                    }
                    if (outputItem.outputClause().outputExpression() != null)
                    {
                        ResolveDefault(outputItem.outputClause().outputExpression().expression());
                    }
                }
            }
        }

        public void Resolve(VBAParser.InputStmtContext context)
        {
            ResolveDefault(context.markedFileNumber().expression());
            foreach (var inputVariable in context.inputList().inputVariable())
            {
                ResolveDefault(inputVariable.expression());
            }
        }

        public void Resolve(VBAParser.PutStmtContext context)
        {
            ResolveFileNumber(context.fileNumber());
            if (context.recordNumber() != null)
            {
                ResolveDefault(context.recordNumber().expression());
            }
            if (context.data() != null)
            {
                ResolveDefault(context.data().expression());
            }
        }

        public void Resolve(VBAParser.GetStmtContext context)
        {
            ResolveFileNumber(context.fileNumber());
            if (context.recordNumber() != null)
            {
                ResolveDefault(context.recordNumber().expression());
            }
            if (context.variable() != null)
            {
                ResolveDefault(context.variable().expression());
            }
        }

        public void Resolve(VBAParser.LsetStmtContext context)
        {
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr);
            }
        }

        public void Resolve(VBAParser.RsetStmtContext context)
        {
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr);
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
                if (context.fieldLength() != null)
                {
                    ResolveDefault(context.fieldLength().expression());
                }
                return;
            }
            ResolveType(asType.complexType().expression());
        }

        public void Resolve(VBAParser.ForNextStmtContext context)
        {
            var firstExpression = _bindingService.ResolveDefault(
                _moduleDeclaration,
                _currentParent,
                context.expression()[0],
                GetInnerMostWithExpression(),
                StatementResolutionContext.Undefined);
            if (firstExpression.Classification == ExpressionClassification.ResolutionFailed)
            {
                _boundExpressionVisitor.AddIdentifierReferences(
                    firstExpression,
                    _qualifiedModuleName,
                    _currentScope,
                    _currentParent);
            }
            else
            {
                // In "For expr1 = expr2" the "expr1 = expr2" part is treated as a single expression.
                var binOp = (BinaryOpExpression)firstExpression;
                var assignmentExpr = binOp.Left;
                var fromExpr = binOp.Right;
                // each iteration counts as an assignment
                _boundExpressionVisitor.AddIdentifierReferences(
                    assignmentExpr,
                    _qualifiedModuleName,
                    _currentScope,
                    _currentParent,
                    true);
                _boundExpressionVisitor.AddIdentifierReferences(
                    assignmentExpr,
                    _qualifiedModuleName,
                    _currentScope,
                    _currentParent);
                _boundExpressionVisitor.AddIdentifierReferences(
                    fromExpr,
                    _qualifiedModuleName,
                    _currentScope,
                    _currentParent);
            }
            for (int exprIndex = 1; exprIndex < context.expression().Count; exprIndex++)
            {
                ResolveDefault(context.expression()[exprIndex]);
            }
        }

        public void Resolve(VBAParser.ForEachStmtContext context)
        {
            var firstExpression = _bindingService.ResolveDefault(
                _moduleDeclaration,
                _currentParent,
                context.expression()[0],
                GetInnerMostWithExpression(),
                StatementResolutionContext.Undefined);
            if (firstExpression.Classification == ExpressionClassification.ResolutionFailed)
            {

                _boundExpressionVisitor.AddIdentifierReferences(
                    firstExpression,
                    _qualifiedModuleName,
                    _currentScope,
                    _currentParent);
            }
            else
            {
                // each iteration counts as an assignment
                _boundExpressionVisitor.AddIdentifierReferences(
                    firstExpression,
                    _qualifiedModuleName,
                    _currentScope,
                    _currentParent,
                    true);
                _boundExpressionVisitor.AddIdentifierReferences(
                    firstExpression,
                    _qualifiedModuleName,
                    _currentScope,
                    _currentParent);
            }
            for (int exprIndex = 1; exprIndex < context.expression().Count; exprIndex++)
            {
                ResolveDefault(context.expression()[exprIndex]);
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
                var identifier = context.identifier().identifierValue().GetText();
                var callee = eventDeclaration;
                callee.AddReference(
                    _qualifiedModuleName,
                    _currentScope,
                    _currentParent,
                    callSiteContext,
                    identifier,
                    callee,
                    callSiteContext.GetSelection(),
                    _annotationService.FindAnnotations(_qualifiedModuleName, callSiteContext.GetSelection().StartLine));
            }
            if (context.eventArgumentList() == null)
            {
                return;
            }
            foreach (var eventArgument in context.eventArgumentList().eventArgument())
            {
                ResolveDefault(eventArgument.expression());
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

        public void Resolve(VBAParser.CircleSpecialFormContext context)
        {
            foreach (var expr in context.expression())
            {
                ResolveDefault(expr);
            }
            ResolveTuple(context.tuple());
        }

        public void Resolve(VBAParser.ScaleSpecialFormContext context)
        {
            if (context.expression() != null)
            {
                ResolveDefault(context.expression());

            }
            foreach (var tuple in context.tuple())
            {
                ResolveTuple(tuple);
            }
        }

        private void ResolveTuple(VBAParser.TupleContext tuple)
        {
            foreach (var expr in tuple.expression())
            {
                ResolveDefault(expr);
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
                    ResolveDefault(enumMember.expression());
                }
            }
        }
    }
}

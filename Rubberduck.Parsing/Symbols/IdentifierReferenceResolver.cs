using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
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
            _boundExpressionVisitor = new BoundExpressionVisitor();
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
            Declaration qualifier = null;
            var expr = context.withStmtExpression();
            var typeExpression = expr.GetText();
            var boundExpression = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent, typeExpression,
                GetInnerMostWithExpression(), ResolutionStatementContext.Undefined);
            if (boundExpression != null)
            {
                _boundExpressionVisitor.AddIdentifierReferences(boundExpression,
                    (exprCtx, identifier, declaration) =>
                        CreateReference(expr, identifier, declaration,
                            RubberduckParserState.CreateBindingSelection(expr, exprCtx)));
                qualifier = boundExpression.ReferencedDeclaration;
            }
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

        private IdentifierReference CreateReference(ParserRuleContext callSiteContext, string identifier,
            Declaration callee, Selection selection, bool isAssignmentTarget = false,
            bool hasExplicitLetStatement = false)
        {
            if (callSiteContext == null || _currentScope == null)
            {
                return null;
            }
            var annotations = FindAnnotations(selection.StartLine);
            return new IdentifierReference(_qualifiedModuleName, _currentScope, _currentParent, identifier, selection,
                callSiteContext, callee, isAssignmentTarget, hasExplicitLetStatement, annotations);
        }

        private IEnumerable<IAnnotation> FindAnnotations(int line)
        {
            var annotations = new List<IAnnotation>();
            var moduleAnnotations = _declarationFinder.ModuleAnnotations(_qualifiedModuleName).ToList();

            // VBE 1-based indexing
            for (var i = line - 1; i >= 1; i--)
            {
                var annotation = moduleAnnotations.SingleOrDefault(a => a.QualifiedSelection.Selection.StartLine == i);

                if (annotation == null)
                {
                    break;
                }

                annotations.Add(annotation);
            }

            return annotations;
        }

        public void Resolve(VBAParser.OnErrorStmtContext context)
        {
            if (context.valueStmt() == null)
            {
                return;
            }
            ResolveLabel(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.ErrorStmtContext context)
        {
            ResolveDefault(context.valueStmt(), context.valueStmt().GetText());
        }

        private void ResolveLabel(ParserRuleContext context, string label)
        {
            var labelDeclaration = _bindingService.ResolveGoTo(_currentParent, label);
            if (labelDeclaration != null)
            {
                labelDeclaration.AddReference(CreateReference(context, context.GetText(), labelDeclaration,
                    context.GetSelection()));
            }
        }

        private void ResolveDefault(ParserRuleContext context, string expression,
            ResolutionStatementContext statementContext = ResolutionStatementContext.Undefined,
            bool isAssignmentTarget = false, bool hasExplicitLetStatement = false)
        {
            var boundExpression = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent, expression,
                GetInnerMostWithExpression(), statementContext);
            if (boundExpression != null)
            {
                _boundExpressionVisitor.AddIdentifierReferences(boundExpression,
                    (exprCtx, identifier, declaration) =>
                        CreateReference(context, identifier, declaration,
                            RubberduckParserState.CreateBindingSelection(context, exprCtx), isAssignmentTarget,
                            hasExplicitLetStatement));
            }
            else
            {
                Debug.WriteLine(
                    string.Format(
                        "Failed to resolve {0}. Possible causes include: COM Coclass/Interface mixup / Alias / Bug in the resolver.",
                        expression));
            }
        }

        private void ResolveType(ParserRuleContext context, string expression)
        {
            var boundExpression = _bindingService.ResolveType(_moduleDeclaration, _currentParent, expression);
            if (boundExpression != null)
            {
                _boundExpressionVisitor.AddIdentifierReferences(boundExpression,
                    (exprCtx, identifier, declaration) =>
                        CreateReference(context, identifier, declaration,
                            RubberduckParserState.CreateBindingSelection(context, exprCtx)));
            }
        }

        public void Resolve(VBAParser.GoToStmtContext context)
        {
            ResolveLabel(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.OnGoToStmtContext context)
        {
            ResolveDefault(context.valueStmt()[0], context.valueStmt()[0].GetText());
            for (int labelIndex = 1; labelIndex < context.valueStmt().Count; labelIndex++)
            {
                ResolveLabel(context.valueStmt()[labelIndex], context.valueStmt()[labelIndex].GetText());
            }
        }

        public void Resolve(VBAParser.GoSubStmtContext context)
        {
            ResolveLabel(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.OnGoSubStmtContext context)
        {
            ResolveDefault(context.valueStmt()[0], context.valueStmt()[0].GetText());
            for (int labelIndex = 1; labelIndex < context.valueStmt().Count; labelIndex++)
            {
                ResolveLabel(context.valueStmt()[labelIndex], context.valueStmt()[labelIndex].GetText());
            }
        }

        public void Resolve(VBAParser.RedimStmtContext context)
        {
            foreach (var redimStmt in context.redimSubStmt())
            {
                foreach (var dimSpec in redimStmt.subscripts().subscript())
                {
                    foreach (var expr in dimSpec.valueStmt())
                    {
                        ResolveDefault(expr, expr.GetText());
                    }
                }
            }
        }

        public void Resolve(VBAParser.WhileWendStmtContext context)
        {
            ResolveDefault(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.DoLoopStmtContext context)
        {
            if (context.valueStmt() == null)
            {
                return;
            }
            ResolveDefault(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.IfStmtContext context)
        {
            ResolveDefault(context.booleanExpression(), context.booleanExpression().GetText());
            if (context.elseIfBlock() != null)
            {
                foreach (var elseIfBlock in context.elseIfBlock())
                {
                    ResolveDefault(elseIfBlock.booleanExpression(), elseIfBlock.booleanExpression().GetText());
                }
            }
        }

        public void Resolve(VBAParser.SingleLineIfStmtContext context)
        {
            // The listOrLabel rule could be resolved separately but since it's such a special case, only appearing in
            // single-line-if-statements, we do it here for better understanding.
            if (context.ifWithEmptyThen() != null)
            {
                ResolveDefault(context.ifWithEmptyThen().booleanExpression(), context.ifWithEmptyThen().booleanExpression().GetText());
                ResolveListOrLabel(context.ifWithEmptyThen().singleLineElseClause().listOrLabel());
            }
            else
            {
                ResolveDefault(context.ifWithNonEmptyThen().booleanExpression(), context.ifWithNonEmptyThen().booleanExpression().GetText());
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
            ResolveDefault(context.valueStmt(), context.valueStmt().GetText());
            if (context.sC_Case() != null)
            {
                foreach (var caseClauseBlock in context.sC_Case())
                {
                    var caseClause = caseClauseBlock.sC_Cond();
                    if (caseClause is VBAParser.CaseCondSelectionContext)
                    {
                        foreach (var selectClause in ((VBAParser.CaseCondSelectionContext)caseClause).sC_Selection())
                        {
                            if (selectClause is VBAParser.CaseCondIsContext)
                            {
                                var ctx = (VBAParser.CaseCondIsContext)selectClause;
                                ResolveDefault(ctx.valueStmt(), ctx.valueStmt().GetText());
                            }
                            else if (selectClause is VBAParser.CaseCondToContext)
                            {
                                var ctx = (VBAParser.CaseCondToContext)selectClause;
                                ResolveDefault(ctx.valueStmt()[0], ctx.valueStmt()[0].GetText());
                                ResolveDefault(ctx.valueStmt()[0], ctx.valueStmt()[0].GetText());
                            }
                            else
                            {
                                var ctx = (VBAParser.CaseCondValueContext)selectClause;
                                ResolveDefault(ctx.valueStmt(), ctx.valueStmt().GetText());
                            }
                        }
                    }
                }
            }
        }

        public void Resolve(VBAParser.LetStmtContext context)
        {
            var letStatement = context.LET();
            ResolveDefault(context.valueStmt()[0], context.valueStmt()[0].GetText(),
                ResolutionStatementContext.LetStatement, true, letStatement != null);
            ResolveDefault(context.valueStmt()[1], context.valueStmt()[1].GetText());
        }

        public void Resolve(VBAParser.SetStmtContext context)
        {
            ResolveDefault(context.valueStmt()[0], context.valueStmt()[0].GetText(),
                ResolutionStatementContext.SetStatement, true, false);
            ResolveDefault(context.valueStmt()[1], context.valueStmt()[1].GetText());
        }

        public void Resolve(VBAParser.ExplicitCallStmtContext context)
        {
            ResolveDefault(context.explicitCallStmtExpression(), context.explicitCallStmtExpression().GetText(),
                ResolutionStatementContext.CallStatement);
        }

        public void Resolve(VBAParser.ConstStmtContext context)
        {
            foreach (var constStmt in context.constSubStmt())
            {
                ResolveDefault(constStmt.valueStmt(), constStmt.valueStmt().GetText());
            }
        }

        public void Resolve(VBAParser.EraseStmtContext context)
        {
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.OpenStmtContext context)
        {
            ResolveDefault(context.pathName(), context.pathName().GetText());
            ResolveDefault(context.fileNumber(), context.fileNumber().GetText());
            if (context.lenClause() != null)
            {
                ResolveDefault(context.lenClause().recLength(), context.lenClause().recLength().GetText());
            }
        }

        public void Resolve(VBAParser.CloseStmtContext context)
        {
            if (context.fileNumberList() != null)
            {
                foreach (var fileNumber in context.fileNumberList().fileNumber())
                {
                    ResolveDefault(fileNumber, fileNumber.GetText());
                }
            }
        }

        public void Resolve(VBAParser.SeekStmtContext context)
        {
            ResolveDefault(context.fileNumber(), context.fileNumber().GetText());
            ResolveDefault(context.position(), context.position().GetText());
        }

        public void Resolve(VBAParser.LockStmtContext context)
        {
            ResolveDefault(context.fileNumber(), context.fileNumber().GetText());
            ResolveRecordRange(context.recordRange());
        }

        public void Resolve(VBAParser.UnlockStmtContext context)
        {
            ResolveDefault(context.fileNumber(), context.fileNumber().GetText());
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
                ResolveDefault(recordRange.startRecordNumber(), recordRange.startRecordNumber().GetText());
            }
            if (recordRange.endRecordNumber() != null)
            {
                ResolveDefault(recordRange.endRecordNumber(), recordRange.endRecordNumber().GetText());
            }
        }

        public void Resolve(VBAParser.LineInputStmtContext context)
        {
            ResolveDefault(context.markedFileNumber(), context.markedFileNumber().GetText());
            ResolveDefault(context.variableName(), context.variableName().GetText());
        }

        public void Resolve(VBAParser.WidthStmtContext context)
        {
            ResolveDefault(context.markedFileNumber(), context.markedFileNumber().GetText());
            ResolveDefault(context.lineWidth(), context.lineWidth().GetText());
        }

        public void Resolve(VBAParser.PrintStmtContext context)
        {
            ResolveDefault(context.markedFileNumber(), context.markedFileNumber().GetText());
            ResolveOutputList(context.outputList());
        }

        public void Resolve(VBAParser.WriteStmtContext context)
        {
            ResolveDefault(context.markedFileNumber(), context.markedFileNumber().GetText());
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
                        ResolveDefault(outputItem.outputClause().spcClause().spcNumber(), outputItem.outputClause().spcClause().spcNumber().GetText());
                    }
                    if (outputItem.outputClause().tabClause() != null && outputItem.outputClause().tabClause().tabNumberClause() != null)
                    {
                        ResolveDefault(outputItem.outputClause().tabClause().tabNumberClause().tabNumber(), outputItem.outputClause().tabClause().tabNumberClause().tabNumber().GetText());
                    }
                    if (outputItem.outputClause().outputExpression() != null)
                    {
                        ResolveDefault(outputItem.outputClause().outputExpression(), outputItem.outputClause().outputExpression().GetText());
                    }
                }
            }
        }

        public void Resolve(VBAParser.InputStmtContext context)
        {
            ResolveDefault(context.markedFileNumber(), context.markedFileNumber().GetText());
            foreach (var inputVariable in context.inputList().inputVariable())
            {
                ResolveDefault(inputVariable, inputVariable.GetText());
            }
        }

        public void Resolve(VBAParser.PutStmtContext context)
        {
            ResolveDefault(context.fileNumber(), context.fileNumber().GetText());
            if (context.recordNumber() != null)
            {
                ResolveDefault(context.recordNumber(), context.recordNumber().GetText());
            }
            if (context.data() != null)
            {
                ResolveDefault(context.data(), context.data().GetText());
            }
        }

        public void Resolve(VBAParser.GetStmtContext context)
        {
            ResolveDefault(context.fileNumber(), context.fileNumber().GetText());
            if (context.recordNumber() != null)
            {
                ResolveDefault(context.recordNumber(), context.recordNumber().GetText());
            }
            if (context.variable() != null)
            {
                ResolveDefault(context.variable(), context.variable().GetText());
            }
        }

        public void Resolve(VBAParser.LsetStmtContext context)
        {
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.RsetStmtContext context)
        {
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.AsTypeClauseContext context)
        {
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
                if (context.fieldLength() != null && context.fieldLength().identifier() != null)
                {
                    ResolveDefault(context.fieldLength().identifier(), context.fieldLength().identifier().GetText());
                }
                return;
            }
            ResolveType(asType.complexType(), asType.complexType().GetText());
        }

        public void Resolve(VBAParser.ForNextStmtContext context)
        {
            var firstExpression = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent,
                context.valueStmt()[0].GetText(), GetInnerMostWithExpression(), ResolutionStatementContext.Undefined);
            if (firstExpression != null)
            {
                // each iteration counts as an assignment
                _boundExpressionVisitor.AddIdentifierReferences(firstExpression,
                    (exprCtx, identifier, declaration) =>
                        CreateReference(context.valueStmt()[0], identifier, declaration,
                            RubberduckParserState.CreateBindingSelection(context.valueStmt()[0], exprCtx), true));
                // each iteration also counts as a plain usage
                _boundExpressionVisitor.AddIdentifierReferences(firstExpression,
                    (exprCtx, identifier, declaration) =>
                        CreateReference(context.valueStmt()[0], identifier, declaration,
                            RubberduckParserState.CreateBindingSelection(context.valueStmt()[0], exprCtx)));
            }
            for (int exprIndex = 1; exprIndex < context.valueStmt().Count; exprIndex++)
            {
                ResolveDefault(context.valueStmt()[exprIndex], context.valueStmt()[exprIndex].GetText());
            }
        }

        public void Resolve(VBAParser.ForEachStmtContext context)
        {
            var firstExpression = _bindingService.ResolveDefault(_moduleDeclaration, _currentParent,
                context.valueStmt()[0].GetText(), GetInnerMostWithExpression(), ResolutionStatementContext.Undefined);
            if (firstExpression != null)
            {
                // each iteration counts as an assignment
                _boundExpressionVisitor.AddIdentifierReferences(firstExpression,
                    (exprCtx, identifier, declaration) =>
                        CreateReference(context.valueStmt()[0], identifier, declaration,
                            RubberduckParserState.CreateBindingSelection(context.valueStmt()[0], exprCtx), true));
                // each iteration also counts as a plain usage
                _boundExpressionVisitor.AddIdentifierReferences(firstExpression,
                    (exprCtx, identifier, declaration) =>
                        CreateReference(context.valueStmt()[0], identifier, declaration,
                            RubberduckParserState.CreateBindingSelection(context.valueStmt()[0], exprCtx)));
            }

            for (int exprIndex = 1; exprIndex < context.valueStmt().Count; exprIndex++)
            {
                ResolveDefault(context.valueStmt()[exprIndex], context.valueStmt()[exprIndex].GetText());
            }
        }

        public void Resolve(VBAParser.ImplementsStmtContext context)
        {
            ResolveType(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.RaiseEventStmtContext context)
        {
            var eventDeclaration = _bindingService.ResolveEvent(_moduleDeclaration, context.identifier().GetText());
            if (eventDeclaration != null)
            {
                eventDeclaration.AddReference(CreateReference(context.identifier(), context.identifier().GetText(),
                    eventDeclaration, context.identifier().GetSelection()));
            }
            ResolveArgsCall(context.argsCall());
        }

        public void Resolve(VBAParser.MidStmtContext context)
        {
            ResolveArgsCall(context.argsCall());
        }

        private void ResolveArgsCall(VBAParser.ArgsCallContext argsCall)
        {
            if (argsCall == null)
            {
                return;
            }
            foreach (var argCall in argsCall.argCall())
            {
                ResolveDefault(argCall.valueStmt(), argCall.valueStmt().GetText());
            }
        }

        public void Resolve(VBAParser.ResumeStmtContext context)
        {
            if (context.valueStmt() == null)
            {
                return;
            }
            ResolveLabel(context.valueStmt(), context.valueStmt().GetText());
        }

        public void Resolve(VBAParser.CircleSpecialFormContext context)
        {
            foreach (var expr in context.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
            ResolveTuple(context.tuple());
        }

        public void Resolve(VBAParser.ScaleSpecialFormContext context)
        {
            if (context.valueStmt() != null)
            {
                ResolveDefault(context.valueStmt(), context.valueStmt().GetText());

            }
            foreach (var tuple in context.tuple())
            {
                ResolveTuple(tuple);
            }
        }

        private void ResolveTuple(VBAParser.TupleContext tuple)
        {
            foreach (var expr in tuple.valueStmt())
            {
                ResolveDefault(expr, expr.GetText());
            }
        }

        public void Resolve(VBAParser.ImplicitCallStmt_InBlockContext context)
        {
            ParserRuleContext subContext;
            if (context.iCS_B_MemberProcedureCall() != null)
            {
                subContext = context.iCS_B_MemberProcedureCall();
            }
            else
            {
                subContext = context.iCS_B_ProcedureCall();
            }
            string expr = subContext.GetText();
            // This represents a CALL statement without the CALL keyword which is slightly different than a normal expression because it does not allow parentheses around its argument list.
            ResolveDefault(subContext, expr, ResolutionStatementContext.CallStatement);
        }

        public void Resolve(VBAParser.EnumerationStmtContext context)
        {
            if (context.enumerationStmt_Constant() == null)
            {
                return;
            }
            foreach (var enumMember in context.enumerationStmt_Constant())
            {
                if (enumMember.valueStmt() != null)
                {
                    ResolveDefault(enumMember.valueStmt(), enumMember.valueStmt().GetText());
                }
            }
        }
    }
}

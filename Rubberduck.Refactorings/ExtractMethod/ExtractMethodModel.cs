using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.Refactorings.Exceptions.ExtractMethod;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodModel : IRefactoringModel
    {
        private List<string> _parametersToExtract;
        private IEnumerable<Declaration> _declarationsToMoveOut;
        private IEnumerable<Declaration> _declarationsToMoveIn;
        private readonly IIndenter _indenter;
        private readonly IExtractedMethod extractedMethod;

        public IEnumerable<ParserRuleContext> SelectedContexts { get; }
        public QualifiedSelection QualifiedSelection { get; }
        public IEnumerable<string> ComponentNames =>
            _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Member).Where(d => d.ComponentName == QualifiedSelection.QualifiedName.ComponentName)
                .Select(d => d.IdentifierName);
        public string SourceMethodName { get => TargetMethod.IdentifierName; }
        public Declaration TargetMethod { get; set; }
        public string NewMethodName
        {
            get => extractedMethod.MethodName;
            set => extractedMethod.NewMethodNameBase = value;
        }

        private ExtractMethodParameter _returnParameter;
        public ExtractMethodParameter ReturnParameter
        {
            get => _returnParameter ?? ExtractMethodParameter.None;
            set => _returnParameter = value ?? ExtractMethodParameter.None;
        }

        public bool ModuleContainsCompilationDirectives { get; set; }

        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public ExtractMethodModel(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<ParserRuleContext> selectedContexts, QualifiedSelection selection, Declaration target, IIndenter indenter)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _indenter = indenter;
            SelectedContexts = selectedContexts;
            QualifiedSelection = selection;
            TargetMethod = target;

            extractedMethod = new ExtractedMethod(declarationFinderProvider);
            Setup();
        }

        private void Setup()
        {
            var functionReturnValueAssignments = TargetMethod.References
                .Where(r => QualifiedSelection.Selection.Contains(r.Selection) &&
                    (r.IsAssignment || r.IsSetAssignment));
            if (functionReturnValueAssignments.Count() != 0)
            {
                var firstSelection = functionReturnValueAssignments.FirstOrDefault().QualifiedSelection;
                var message = RefactoringsUI.ExtractMethod_InvalidMessageSelectionModifiesParentFunctionReturn;
                throw new InvalidTargetSelectionException(firstSelection, message);
            }

            var declarationsInSelection = GetDeclarationsInSelection(QualifiedSelection);

            var sourceMethodParameters = ((IParameterizedDeclaration)TargetMethod).Parameters;
            var sourceMethodSelection = new QualifiedSelection(QualifiedSelection.QualifiedName,
                new Selection(TargetMethod.Context.Start.Line, TargetMethod.Context.Start.Column, TargetMethod.Context.Stop.Line, TargetMethod.Context.Stop.Column));

            var declarationsInParentMethod = GetDeclarationsInSelection(sourceMethodSelection);

            //List of "inbound" variables. Parent procedure parameters + explicit dims which get referenced inside the selection.
            var inboundParameters = sourceMethodParameters.Where(d => d.References.Any(r => QualifiedSelection.Selection.Contains(r.Selection)));
            var inboundLocalVariables = declarationsInParentMethod
                .Where(d => d.References.Any(r => QualifiedSelection.Selection.Contains(r.Selection)) &&
                       (d.References.Any(r => r.Selection.EndLine < QualifiedSelection.Selection.StartLine ||
                                        (r.Selection.EndLine == QualifiedSelection.Selection.StartLine &&
                                         r.Selection.EndColumn < QualifiedSelection.Selection.StartColumn))));
            var inboundVariables = inboundParameters.Concat(inboundLocalVariables);

            //List of "outbound" variables. Any variables assigned a value in the selection which are then referenced after the selection
            var outboundVariables = sourceMethodParameters.Concat(declarationsInParentMethod)
                .Where(d => d.References.Any(r => QualifiedSelection.Selection.Contains(r.Selection)) &&
                                (d.References.Any(r => r.Selection.StartLine > QualifiedSelection.Selection.EndLine ||
                                                 (r.Selection.StartLine == QualifiedSelection.Selection.EndLine &&
                                                  r.Selection.StartColumn > QualifiedSelection.Selection.EndColumn))));

            SetUpParameters(inboundVariables, outboundVariables);

            //Variables to have declarations moved out of the selection
            // - where declaration is in the selection and it is a ByRef variable i.e. intersection of declarations in selection and outbound
            _declarationsToMoveOut = declarationsInSelection.Intersect(outboundVariables)
                .OrderByDescending(d => d.Selection.StartLine)
                .ThenByDescending(d => d.Selection.StartColumn);

            //Variables to have declarations moved into the selection
            // - where declaration is before the selection but only references are inside the selection
            _declarationsToMoveIn = declarationsInParentMethod.Except(declarationsInSelection)
                                    .Where(d => d.References.Any(r => QualifiedSelection.Selection.Contains(r.Selection)) &&
                                           !d.References.Any(r => r.Selection.EndLine < QualifiedSelection.Selection.StartLine ||
                                                            (r.Selection.EndLine == QualifiedSelection.Selection.StartLine &&
                                                             r.Selection.EndColumn < QualifiedSelection.Selection.StartColumn)) &&
                                           !d.References.Any(r => r.Selection.StartLine > QualifiedSelection.Selection.EndLine ||
                                                            (r.Selection.StartLine == QualifiedSelection.Selection.EndLine &&
                                                             r.Selection.StartColumn > QualifiedSelection.Selection.EndColumn)));
        }

        private IOrderedEnumerable<Declaration> GetDeclarationsInSelection(QualifiedSelection qualifiedSelection)
        {
            //Had to add check for declaration type name is VariableDeclaration despite already filtering on 
            //DeclarationType.Variable. This is due to Debug.Print being picked up for some reason!
            return _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(d => qualifiedSelection.Selection.Contains(d.Selection) &&
                            d.QualifiedName.QualifiedModuleName == qualifiedSelection.QualifiedName &&
                            d.GetType().Name == "VariableDeclaration")
                .OrderBy(d => d.Selection.StartLine)
                .ThenBy(d => d.Selection.StartColumn);
        }

        private void SetUpParameters(IEnumerable<Declaration> inboundVariables, IEnumerable<Declaration> outboundVariables)
        {
            Parameters = new ObservableCollection<ExtractMethodParameter>();

            foreach (var declaration in inboundVariables.Union(outboundVariables))
            {
                ExtractMethodParameterType paramType;
                bool canReturn;
                if (inboundVariables.Contains(declaration) && !outboundVariables.Contains(declaration))
                {
                    //List of "inbound" only variables (to be passed ByVal unless array type)
                    if (declaration.IsArray)
                    {
                        paramType = ExtractMethodParameterType.ByRefParameter;
                        canReturn = false;
                    }
                    else
                    {
                        paramType = ExtractMethodParameterType.ByValParameter;
                        canReturn = false;
                    }
                }
                else if (inboundVariables.Contains(declaration))
                {
                    //List of "inbound" and "outbound" variables (to be passed ByRef)
                    paramType = ExtractMethodParameterType.ByRefParameter;
                    canReturn = false;
                }
                else
                {
                    //List of "outbound" only variables (to be passed ByRef OR set as a return value)
                    paramType = ExtractMethodParameterType.ByRefParameter;
                    canReturn = true;
                }
                Parameters.Add(new ExtractMethodParameter(declaration, paramType, canReturn));
            }
        }

        private int SelectionIndentation;

        //Code excluding declarations that are to be moved out of the selection
        public string SelectedCodeToExtract 
        {
            get
            {
                var targetMethodCode = TargetMethod.Context.GetText().Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                var targetMethodSelection = TargetMethod.Selection;
                var selectionToExtract = QualifiedSelection.Selection;
                var selectionCode = targetMethodCode.Skip(selectionToExtract.StartLine - targetMethodSelection.StartLine)
                                                    .Take(selectionToExtract.EndLine - selectionToExtract.StartLine + 1)
                                                    .ToList();
                SelectionIndentation = selectionCode[0].Length - selectionCode[0].TrimStart(' ').Length;

                //Remove any code on the first and last lines of selection that isn't inside the selection
                var firstLineStartExcludedTo = selectionToExtract.StartColumn - 1;
                var firstLineWasPruned = false;
                if (selectionToExtract.StartColumn > 1)
                {
                    selectionCode[0] = selectionCode[0].Substring(firstLineStartExcludedTo);
                    firstLineWasPruned = true;
                }
                var lastLineEndExcludedFrom = selectionToExtract.EndColumn - 1;
                if (firstLineWasPruned && selectionToExtract.StartLine == selectionToExtract.EndLine)
                {
                    lastLineEndExcludedFrom -= firstLineStartExcludedTo;
                }
                var lastLineWasPruned = false;
                if (selectionToExtract.EndColumn < targetMethodCode[selectionToExtract.EndLine - targetMethodSelection.StartLine].Length)
                {
                    selectionCode[selectionCode.Count - 1] = selectionCode[selectionCode.Count - 1].Substring(0, lastLineEndExcludedFrom);
                    lastLineWasPruned = true;
                }
                //Remove code for declarations that need to be moved out of the selection to stay in the parent method
                foreach (var decl in _declarationsToMoveOut.Except(_declarationsToMoveOut.Where(d => d.IdentifierName == ReturnParameter.Name)))
                {
                    var variableListStmt = (VBAParser.VariableListStmtContext)decl.Context.Parent;
                    //TODO - Build ability to cope with multiple variables in a list declaration.
                    //Maybe just blank the whole line and rewrite any declarations deleted by iterating over those not in the list to move?
                    var numVariablesDeclared = variableListStmt.ChildCount;
                    if (numVariablesDeclared > 1)
                    {
                        throw new UnableToMoveVariableDeclarationException(decl);
                    }

                    var fullDeclStmt = (VBAParser.VariableStmtContext)decl.Context.Parent.Parent;

                    //Check that declaration not split over multiple lines (until able to handle this)
                    var startLine = fullDeclStmt.Start.Line;
                    var endLine = fullDeclStmt.Stop.Line;
                    if (endLine > startLine)
                    {
                        throw new UnableToMoveVariableDeclarationException(decl);
                    }

                    int declLineStartPrunedAmount = (startLine == selectionToExtract.StartLine && firstLineWasPruned) ? firstLineStartExcludedTo : 0;
                    int declLineEndPrunedAmount = (startLine == selectionToExtract.EndLine && lastLineWasPruned) ? lastLineEndExcludedFrom : 0;
                    //Remove declaration range from selected code
                    RemoveDeclaration(decl, selectionCode, startLine - selectionToExtract.StartLine, declLineStartPrunedAmount, declLineEndPrunedAmount);
                }

                return string.Join(Environment.NewLine, selectionCode);
            }
        }

        private void RemoveDeclaration(Declaration decl, List<string> selectionCode, int declCodeIndex, int declLineStartPrunedAmount, int declLineEndPrunedAmount)
        {
            //Search for other BlockStatements on the same line if need to just cut out the declaration
            string leftPart = string.Empty;
            string rightPart = string.Empty;
            int cutFrom = 0;
            int cutTo = selectionCode[declCodeIndex].Length;
            var declBlockStmtContext = decl.Context.GetAncestor<VBAParser.BlockStmtContext>();
            VBAParser.EndOfStatementContext precedingContext;
            VBAParser.BlockStmtContext precedingBlockStmtContext;
            bool hasPrior = false;
            if (declBlockStmtContext.TryGetPrecedingContext(out precedingContext))
            {
                if (precedingContext.TryGetPrecedingContext(out precedingBlockStmtContext))
                {
                    if (precedingBlockStmtContext.Stop.Line == declBlockStmtContext.Start.Line &&
                        QualifiedSelection.Contains(new QualifiedSelection(QualifiedSelection.QualifiedName, precedingBlockStmtContext.GetSelection())))
                    {
                        hasPrior = true;
                        //start of cut to exclude end of statement (i.e. colon separator) following previous block statement
                        cutFrom = precedingBlockStmtContext.Stop.EndColumn();
                        leftPart = selectionCode[declCodeIndex].Substring(0, cutFrom);
                    }
                }
            }

            VBAParser.EndOfStatementContext followingContext;
            VBAParser.BlockStmtContext followingBlockStmtContext;
            bool hasFollower = false;
            if (declBlockStmtContext.TryGetFollowingContext(out followingContext))
            {
                if (followingContext.TryGetFollowingContext(out followingBlockStmtContext))
                {
                    if (followingBlockStmtContext.Start.Line == declBlockStmtContext.Stop.Line &&
                        QualifiedSelection.Contains(new QualifiedSelection(QualifiedSelection.QualifiedName, followingBlockStmtContext.GetSelection())))
                    {
                        hasFollower = true;
                        //end of cut to just go up to end of block statement to be cut, preserving following end of statement
                        //unless the declaration to move was the first block statement of the line and so we need to remove
                        //the end of statement context too to avoid starting a line with a colon
                        if (hasPrior)
                        {
                            cutTo = declBlockStmtContext.Stop.EndColumn() - declLineStartPrunedAmount;
                        }
                        else
                        {
                            cutTo = followingContext.Stop.EndColumn() - declLineStartPrunedAmount;
                        }
                        rightPart = selectionCode[declCodeIndex].Substring(cutTo, selectionCode[declCodeIndex].Length - cutTo);
                    }
                }
            }

            if (hasFollower || hasPrior)
            {
                selectionCode[declCodeIndex] = leftPart + rightPart;
            }
            else
            {
                //No other code found, remove whole line
                selectionCode.RemoveAt(declCodeIndex);
            }
        }

        public ObservableCollection<ExtractMethodParameter> Parameters { get; set; }

        private IEnumerable<ExtractMethodParameter> ArgumentsToPass => from p in Parameters where p != ReturnParameter select p;

        public string ReplacementCode
        {
            get
            {
                var strings = new List<string> { string.Empty };
                var indentation = new string(' ', SelectionIndentation);

                foreach (var dec in _declarationsToMoveOut)
                {
                    var fullDec = (VBAParser.VariableStmtContext)dec.Context.Parent.Parent;
                    var subscripts = dec.Context.GetDescendent<VBAParser.BoundsListContext>()?.GetText() ?? string.Empty;
                    var identifier = dec.IsArray ? $"{dec.IdentifierName}({subscripts})" : dec.IdentifierName;
                    var declarationType = IsStatic(dec) ? Tokens.Static : Tokens.Dim;
                    var newVariable = dec.AsTypeContext is null
                        ? $"{declarationType} {identifier} {Tokens.As} {Tokens.Variant}"
                        : $"{declarationType} {identifier} {Tokens.As} {(dec.IsSelfAssigned ? Tokens.New + " " : string.Empty)}{dec.AsTypeNameWithoutArrayDesignator}";

                    strings.Add(indentation + newVariable);
                }

                // Make call to new method
                var argList = string.Join(", ", from p in ArgumentsToPass select p.Name);

                if (ReturnParameter == ExtractMethodParameter.None)
                {
                    strings.Add($"{indentation}{NewMethodName} {argList}".TrimEnd());
                }
                else if (ReturnParameter.IsObject)
                {
                    strings.Add($"{indentation}{Tokens.Set} {ReturnParameter.Name} = {NewMethodName}({argList})");
                }
                else
                {
                    strings.Add($"{indentation}{ReturnParameter.Name} = {NewMethodName}({argList})");
                }
                return string.Join(Environment.NewLine, strings);
            }
        }
        public string NewMethodCode
        {
            get
            {
                _parametersToExtract = new List<string>();

                _parametersToExtract.AddRange(from p in ArgumentsToPass select p.ToString());

                var isFunction = ReturnParameter != ExtractMethodParameter.None;

                var strings = new List<string>();
                var returnType = string.Empty;
                if (isFunction)
                {
                    returnType = string.Concat(Tokens.As, " ",
                        ReturnParameter.ToString(ExtractMethodParameterFormat.ReturnDeclaration) ?? Tokens.Variant);
                }
                strings.Add(
                    $@"{Tokens.Private} {(isFunction ? Tokens.Function : Tokens.Sub)} {
                            NewMethodName ?? RefactoringsUI.ExtractMethod_DefaultNewMethodName
                        }({string.Join(", ", _parametersToExtract)}) {returnType}");
                foreach (var dec in _declarationsToMoveIn)
                {
                    strings.Add(dec.Context.GetAncestor<VBAParser.VariableStmtContext>().GetText());
                    //TODO - handle case of variable list having multiple parts (if not excluded by validator)
                }
                if (isFunction)
                {
                    if (!QualifiedSelection.Selection.Contains(ReturnParameter.Declaration.Context))
                    {
                        strings.Add(ReturnParameter.Declaration.Context.GetAncestor<VBAParser.VariableStmtContext>().GetText());
                    }
                }
                strings.AddRange(SelectedCodeToExtract.Split(new[] { Environment.NewLine }, StringSplitOptions.None));
                if (isFunction)
                {
                    strings.Add(string.Empty);
                    string setSection = ReturnParameter.IsObject ? $"{Tokens.Set} " : string.Empty;
                    strings.Add($"{setSection}{NewMethodName} = {ReturnParameter.Name}");
                }
                strings.Add($"{Tokens.End} {(isFunction ? Tokens.Function : Tokens.Sub)}");
                //Add empty strings so have a space after original function
                var indentedStrings = new List<string> { string.Empty }.Concat(_indenter.Indent(strings));

                return string.Join(Environment.NewLine, indentedStrings);
            }
        }

        private static bool IsStatic(Declaration declaration)
        {
            var ctxt = declaration.Context.GetAncestor<VBAParser.VariableStmtContext>();
            if (ctxt?.STATIC() != null)
            {
                return true;
            }

            switch (declaration.ParentDeclaration.Context)
            {
                case VBAParser.FunctionStmtContext func:
                    return func.STATIC() != null;
                case VBAParser.SubStmtContext sub:
                    return sub.STATIC() != null;
                case VBAParser.PropertyLetStmtContext let:
                    return let.STATIC() != null;
                case VBAParser.PropertySetStmtContext set:
                    return set.STATIC() != null;
                case VBAParser.PropertyGetStmtContext get:
                    return get.STATIC() != null;
                default:
                    return false;
            }
        }
    }
}
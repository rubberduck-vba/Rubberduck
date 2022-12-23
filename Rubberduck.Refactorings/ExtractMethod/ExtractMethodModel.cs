using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.Refactorings.Exceptions.ExtractMethod;
using Rubberduck.Parsing;
using System.Text.RegularExpressions;

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
        public string SourceMethodName { get; private set; }
        public Declaration TargetMethod { get; set; }
        //public IEnumerable<Declaration> SourceVariables { get; private set; }
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
            //if (string.IsNullOrWhiteSpace(NewMethodName))
            //{
            //    NewMethodName = RefactoringsUI.ExtractMethod_DefaultNewMethodName; //Check for conflicts - see other document for example code
            //}

            SelectedCode = string.Join(Environment.NewLine, SelectedContexts.Select(c => c.GetText()));

            //SourceVariables = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
            //    .Where(d => (Selection.Selection.Contains(d.Selection) &&
            //                 d.QualifiedName.QualifiedModuleName == Selection.QualifiedName) ||
            //                d.References.Any(r =>
            //                    r.QualifiedModuleName.ComponentName == Selection.QualifiedName.ComponentName
            //                    && r.QualifiedModuleName.ComponentName ==
            //                    d.QualifiedName.QualifiedModuleName.ComponentName
            //                    && Selection.Selection.Contains(r.Selection)))
            //    .OrderBy(d => d.Selection.StartLine)
            //    .ThenBy(d => d.Selection.StartColumn);

            var declarationsInSelection = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(d => QualifiedSelection.Selection.Contains(d.Selection) &&
                            d.QualifiedName.QualifiedModuleName == QualifiedSelection.QualifiedName)
                .OrderBy(d => d.Selection.StartLine)
                .ThenBy(d => d.Selection.StartColumn);
            var referencesOfDeclarationsInSelection = (from dec in declarationsInSelection select dec.References).ToArray(); //debug purposes

            var sourceMethodParameters = ((IParameterizedDeclaration)TargetMethod).Parameters;
            var sourceMethodSelection = new QualifiedSelection(QualifiedSelection.QualifiedName,
                new Selection(TargetMethod.Context.Start.Line, TargetMethod.Context.Start.Column, TargetMethod.Context.Stop.Line, TargetMethod.Context.Stop.Column));

            //TODO - refactor below and declarationsInSelection to generic declarations in selection method if stay consistent
            var declarationsInParentMethod = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(d => sourceMethodSelection.Selection.Contains(d.Selection) &&
                            d.QualifiedName.QualifiedModuleName == sourceMethodSelection.QualifiedName)
                .OrderBy(d => d.Selection.StartLine)
                .ThenBy(d => d.Selection.StartColumn);

            //List of "inbound" variables. Parent procedure parameters + explicit dims which get referenced inside the selection.
            //Ideally excluding those declared but not assigned before the selection. Refinement to change this later
            //No need to check if reference is outside of method because just dealing with local parameters and declarations
            //Add function reference if it is a function?
            //Would ideally identify variables assigned before the selection, not just declared. Could assume any reference before the selection is an assignment?
            //TODO - add case where reference earlier in the same line as where the selection starts (unusual but could exist if using colons to separate multiple statements)
            var inboundParameters = sourceMethodParameters.Where(d => d.References.Any(r => QualifiedSelection.Selection.Contains(r.Selection)));
            var inboundLocalVariables = declarationsInParentMethod.Where(d => d.References.Any(r => QualifiedSelection.Selection.Contains(r.Selection)) &&
                                                                              d.References.Any(r => r.Selection.EndLine < QualifiedSelection.Selection.StartLine));
            var inboundVariables = inboundParameters.Concat(inboundLocalVariables);

            //List of "outbound" variables. Any variables assigned a value in the selection which are then referenced after the selection
            var outboundVariables = sourceMethodParameters.Concat(declarationsInParentMethod)
                .Where(d => d.References.Any(r => QualifiedSelection.Selection.Contains(r.Selection)) && d.References.Any(r => r.Selection.StartLine > QualifiedSelection.Selection.EndLine));

            //Set up parameters
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
                Parameters.Add(new ExtractMethodParameter(declaration.AsTypeNameWithoutArrayDesignator,
                                                          paramType, declaration.IdentifierName,
                                                          declaration.IsArray, declaration.IsObject, canReturn));
            }

            //Variables to have declarations moved out of the selection
            // - where declaration is in the selection and it is a ByRef variable i.e. intersection of declarations in selection and outbound
            _declarationsToMoveOut = declarationsInSelection.Intersect(outboundVariables)
                .OrderByDescending(d => d.Selection.StartLine)
                .ThenByDescending(d => d.Selection.StartColumn);

            //Variables to have declarations moved into the selection
            // - where declaration is before the selection but only references are inside the selection
            _declarationsToMoveIn = declarationsInParentMethod.Except(declarationsInSelection)
                                    .Where(d => d.References.Any(r => QualifiedSelection.Selection.Contains(r.Selection)) &&
                                           !d.References.Any(r => r.Selection.EndLine < QualifiedSelection.Selection.StartLine) &&
                                           !d.References.Any(r => r.Selection.StartLine > QualifiedSelection.Selection.EndLine));

            //List of neither "inbound" or "outbound" (need the declaration copied inside the selection OR moved if careful but can leave inspections to pick up unnecessary declarations)
            //Only applies if the list of inbound variables excludes those declared but not assigned before the selection
            //????

        }

        public string SelectedCode { get; private set; }

        //Code excluding declarations that are to be moved out of the selection
        public string SelectedCodeToExtract 
        {
            get
            {
                //TODO *** ANY WAY TO USE MOVE CLOSER TO USAGE REFACTORING CODE??? ***

                var targetMethodCode = TargetMethod.Context.GetText().Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                //var originalCode = string.Join(Environment.NewLine, SelectedContexts.Select(c => c.GetText()));
                //TODO - create way to map selection to the string or confirm that commented out code above works (including whitespace)

                var targetMethodSelection = TargetMethod.Selection;
                var selectionCode = targetMethodCode.Skip(QualifiedSelection.Selection.StartLine - targetMethodSelection.StartLine)
                                                    .Take(QualifiedSelection.Selection.EndLine - QualifiedSelection.Selection.StartLine + 1)
                                                    .ToList();
                var selectionToExtract = QualifiedSelection.Selection;

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

                    //TODO - Check if multiple block statements on the same line
                    //1) Go up to Block ancester
                    //2) Find BlockStmt enclosing the declaration
                    //3) Check if preceding or following BlockStmts (if exist) are on the same line (preceding.Stop or following.Start)
                    //4) Throw error or rebuild line


                    //Remove declaration range from selected code
                    selectionCode.RemoveAt(startLine - selectionToExtract.StartLine);
                }

                //TODO - Handle first and last lines of selection. Could be affected by previous moves, so need better approach using parse tree.
                //As a start, detect if code on those lines that is not in the selection and throw error
                //if (selectionToExtract.StartColumn > 1)
                //{
                //    selectionCode[0] = selectionCode[0].Substring(selectionToExtract.StartColumn - 1);
                //}
                //if (selectionToExtract.EndColumn < targetMethodCode[selectionToExtract.EndLine - selectionToExtract.StartLine].Length)
                //{
                //    selectionCode[selectionCode.Count - 1] = selectionCode[selectionCode.Count - 1].Substring(0, selectionToExtract.EndColumn);
                //}

                //var selectionsToMoveOut = (from dec in _declarationsToMoveOut 
                //                           orderby dec.Selection.StartLine, dec.Selection.StartColumn 
                //                           select dec.Selection);
                return string.Join(Environment.NewLine, selectionCode);
            }
        }

        public ObservableCollection<ExtractMethodParameter> Parameters { get; set; }

        private IEnumerable<ExtractMethodParameter> ArgumentsToPass => from p in Parameters where p != ReturnParameter select p;

        public string ReplacementCode
        {
            get
            {
                var strings = new List<string> { string.Empty };
                string indentation;
                if (SelectedContexts.First().GetType() == typeof(VBAParser.BlockStmtContext))
                {
                    indentation = FrontPadding((VBAParser.BlockStmtContext)SelectedContexts.First());
                }
                else
                {
                    var enclosingStatementContext = SelectedContexts.First().GetAncestor<VBAParser.BlockStmtContext>();
                    indentation = FrontPadding(enclosingStatementContext);                    
                }

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
                var argList = string.Join(",", from p in ArgumentsToPass select p.Name);

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
                    //go from variableSubStmt to variableListStmt to variableStmt
                    strings.Add(dec.Context.Parent.Parent.GetText());
                    //TODO - handle case of variable list having multiple parts (if not excluded by validator)
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

        private string FrontPadding(VBAParser.BlockStmtContext context)
        {
            var paddingChars = context.Start.Column;
            if (paddingChars > 0)
            {
                return new string(' ', paddingChars);
            }
            else
            {
                return string.Empty;
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
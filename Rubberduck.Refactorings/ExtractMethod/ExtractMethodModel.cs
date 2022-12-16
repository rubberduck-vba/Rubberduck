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

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodModel : IRefactoringModel
    {
        private List<string> _parametersToExtract;
        private IEnumerable<Declaration> _declarationsToMoveOut;
        private IEnumerable<Declaration> _declarationsToMoveIn;
        private readonly IIndenter _indenter;

        public IEnumerable<ParserRuleContext> SelectedContexts { get; }
        public QualifiedSelection Selection { get; }
        public IEnumerable<string> ComponentNames =>
            _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Member).Where(d => d.ComponentName == Selection.QualifiedName.ComponentName)
                .Select(d => d.IdentifierName);
        public string SourceMethodName { get; private set; }
        public Declaration TargetMethod { get; set; }
        //public IEnumerable<Declaration> SourceVariables { get; private set; }
        public string NewMethodName { get; set; }

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
            Selection = selection;
            TargetMethod = target;
            Setup();
        }

        private void Setup()
        {
            ParserRuleContext stmtContext = null;
            stmtContext = GetStatementContext(stmtContext);

            if (string.IsNullOrWhiteSpace(NewMethodName))
            {
                NewMethodName = RefactoringsUI.ExtractMethod_DefaultNewMethodName; //Check for conflicts - see other document for example code
            }

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
                .Where(d => Selection.Selection.Contains(d.Selection) &&
                            d.QualifiedName.QualifiedModuleName == Selection.QualifiedName)
                .OrderBy(d => d.Selection.StartLine)
                .ThenBy(d => d.Selection.StartColumn);
            var referencesOfDeclarationsInSelection = (from dec in declarationsInSelection select dec.References).ToArray(); //debug purposes

            var sourceMethodParameters = ((IParameterizedDeclaration)TargetMethod).Parameters;
            var sourceMethodSelection = new QualifiedSelection(Selection.QualifiedName,
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
            var inboundParameters = sourceMethodParameters.Where(d => d.References.Any(r => Selection.Selection.Contains(r.Selection)));
            var inboundLocalVariables = declarationsInParentMethod.Where(d => d.References.Any(r => Selection.Selection.Contains(r.Selection)) &&
                                                                              d.References.Any(r => r.Selection.EndLine < Selection.Selection.StartLine));
            var inboundVariables = inboundParameters.Concat(inboundLocalVariables);

            //List of "outbound" variables. Any variables assigned a value in the selection which are then referenced after the selection
            var outboundVariables = sourceMethodParameters.Concat(declarationsInParentMethod)
                .Where(d => d.References.Any(r => Selection.Selection.Contains(r.Selection)) && d.References.Any(r => r.Selection.StartLine > Selection.Selection.EndLine));

            //Set up parameters
            Parameters = new ObservableCollection<ExtractMethodParameter>();

            foreach (var declaration in inboundVariables.Union(outboundVariables))
            {
                if (inboundVariables.Contains(declaration) && !outboundVariables.Contains(declaration))
                {
                    //List of "inbound" only variables (to be passed ByVal)
                    Parameters.Add(new ExtractMethodParameter(declaration.AsTypeNameWithoutArrayDesignator,
                                                              ExtractMethodParameterType.ByValParameter,
                                                              declaration.IdentifierName, declaration.IsArray, false));
                }
                else if (inboundVariables.Contains(declaration))
                {
                    //List of "inbound" and "outbound" variables (to be passed ByRef)
                    Parameters.Add(new ExtractMethodParameter(declaration.AsTypeNameWithoutArrayDesignator,
                                                              ExtractMethodParameterType.ByRefParameter,
                                                              declaration.IdentifierName, declaration.IsArray, false));
                }
                else
                {
                    //List of "outbound" only variables (to be passed ByRef OR set as a return value)
                    Parameters.Add(new ExtractMethodParameter(declaration.AsTypeNameWithoutArrayDesignator,
                                                              ExtractMethodParameterType.ByRefParameter,
                                                              declaration.IdentifierName, declaration.IsArray, true));
                }
            }

            //Variables to have declarations moved out of the selection
            // - where declaration is in the selection and it is a ByRef variable i.e. intersection of declarations in selection and outbound
            _declarationsToMoveOut = declarationsInSelection.Intersect(outboundVariables);

            //Variables to have declarations moved into the selection
            // - where declaration is before the selection but only references are inside the selection
            _declarationsToMoveIn = declarationsInParentMethod.Except(declarationsInSelection)
                                    .Where(d => d.References.Any(r => Selection.Selection.Contains(r.Selection)) &&
                                           !d.References.Any(r => r.Selection.EndLine < Selection.Selection.StartLine) &&
                                           !d.References.Any(r => r.Selection.StartLine > Selection.Selection.EndLine));

            //List of neither "inbound" or "outbound" (need the declaration copied inside the selection OR moved if careful but can leave inspections to pick up unnecessary declarations)
            //Only applies if the list of inbound variables excludes those declared but not assigned before the selection
            SelectedCodeToExtract = string.Join(Environment.NewLine, SelectedContexts.Except(from dec in _declarationsToMoveOut select dec.Context).Select(c => c.GetText()));
        }

        private ParserRuleContext GetStatementContext(ParserRuleContext stmtContext)
        {
            var topContext = SelectedContexts.First();
            var currentContext = (RuleContext)topContext;
            do
            {
                switch (currentContext)
                {
                    case VBAParser.FunctionStmtContext stmt:
                        stmtContext = stmt;
                        SourceMethodName = stmt.functionName().GetText();
                        break;
                    case VBAParser.SubStmtContext stmt:
                        stmtContext = stmt;
                        SourceMethodName = stmt.subroutineName().GetText();
                        break;
                    case VBAParser.PropertyGetStmtContext stmt:
                        stmtContext = stmt;
                        SourceMethodName = stmt.functionName().GetText();
                        break;
                    case VBAParser.PropertyLetStmtContext stmt:
                        stmtContext = stmt;
                        SourceMethodName = stmt.subroutineName().GetText();
                        break;
                    case VBAParser.PropertySetStmtContext stmt:
                        stmtContext = stmt;
                        SourceMethodName = stmt.subroutineName().GetText();
                        break;
                }
                currentContext = currentContext.Parent;
            }
            while (currentContext != null && stmtContext == null);
            return stmtContext;
        }

        public string SelectedCode { get; private set; }

        //Code excluding declarations that are to be moved out of the selection
        public string SelectedCodeToExtract { get; private set; }

        public ObservableCollection<ExtractMethodParameter> Parameters { get; set; }

        private IEnumerable<ExtractMethodParameter> ArgumentsToPass => from p in Parameters where p != ReturnParameter select p;

        public string ReplacementCode
        {
            get
            {
                var strings = new List<string>();

                //TODO - find indentation of first line of selection and add same number of spaces to all lines added after the first

                foreach (var dec in _declarationsToMoveOut)
                {
                    //go from variableSubStmt to variableListStmt to variableStmt
                    strings.Add(dec.Context.Parent.Parent.GetText());
                    //TODO - handle case of variable list having multiple parts (if not excluded by validator)
                }

                // Make call to new method
                var argList = string.Join(",", from p in ArgumentsToPass select p.Name);

                if (ReturnParameter == ExtractMethodParameter.None)
                {
                    strings.Add($"{NewMethodName} {argList}");
                }
                else
                {
                    //TODO - make sure can handle setting of objects if return type is an object
                    strings.Add($"{ReturnParameter.Name} = {NewMethodName}({argList})");
                }

                return string.Join(Environment.NewLine, strings);
            }
        }
        public string PreviewCode
        {
            get
            {
                _parametersToExtract = new List<string>();

                _parametersToExtract.AddRange(from p in ArgumentsToPass select p.ToString());

                var isFunction = ReturnParameter != ExtractMethodParameter.None;

                var strings = new List<string>();
                // string.Empty are used to create blank lines as the joins will create a newline each line.
                strings.Add(string.Empty);
                strings.Add(string.Empty);
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
                strings.AddRange(SelectedCodeToExtract.Split(new[] {Environment.NewLine}, StringSplitOptions.None));
                if (isFunction)
                {
                    strings.Add(string.Empty);
                    strings.Add($"{NewMethodName} = {ReturnParameter.Name}");
                }
                strings.Add($"{Tokens.End} {(isFunction ? Tokens.Function : Tokens.Sub)}");
                return string.Join(Environment.NewLine, _indenter.Indent(strings));
            }
        }
    }
}
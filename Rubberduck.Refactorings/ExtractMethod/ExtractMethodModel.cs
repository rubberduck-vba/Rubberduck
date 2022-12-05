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
using Rubberduck.UI;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodModel : IRefactoringModel
    {
        private List<string> _fieldsToExtract;
        private List<string> _parametersToExtract;
        private List<string> _variablesToExtract;

        public IEnumerable<ParserRuleContext> SelectedContexts { get; }
        public QualifiedSelection Selection { get; }

        public string SourceMethodName { get; private set; }
        public Declaration TargetMethod { get; set; }
        public IEnumerable<Declaration> SourceVariables { get; private set; }
        public string NewMethodName { get; set; }

        private ExtractMethodParameter _returnParameter;
        public ExtractMethodParameter ReturnParameter
        {
            get => _returnParameter ?? ExtractMethodParameter.None;
            set => _returnParameter = value ?? ExtractMethodParameter.None;
        }

        //public bool ModuleContainsCompilationDirectives { get; private set; }
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        public ExtractMethodModel(IDeclarationFinderProvider declarationFinderProvider, IEnumerable<ParserRuleContext> selectedContexts, QualifiedSelection selection, Declaration target)
        {
            _declarationFinderProvider = declarationFinderProvider;
            SelectedContexts = selectedContexts;
            Selection = selection;
            TargetMethod = target;
            Setup();
        }

        private void Setup()
        {
            var topContext = SelectedContexts.First();
            ParserRuleContext stmtContext = null;
            var currentContext = (RuleContext)topContext;
            do {
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
            while (currentContext != null && stmtContext == null) ;

            if (string.IsNullOrWhiteSpace(NewMethodName))
            {
                NewMethodName = "ExtractMethod_DefaultNewMethodName"; //RefactoringsUI.ExtractMethod_DefaultNewMethodName
            }

            SelectedCode = string.Join(Environment.NewLine, SelectedContexts.Select(c => c.GetText()));

            //ModuleContainsCompilationDirectives = CodeModule.ContainsCompilationDirectives();

            SourceVariables = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(d => ((Selection.Selection.Contains(d.Selection) &&
                             d.QualifiedName.QualifiedModuleName == Selection.QualifiedName)) ||
                            d.References.Any(r =>
                                r.QualifiedModuleName.ComponentName == Selection.QualifiedName.ComponentName
                                && r.QualifiedModuleName.ComponentName ==
                                d.QualifiedName.QualifiedModuleName.ComponentName
                                && Selection.Selection.Contains(r.Selection)))
                .OrderBy(d => d.Selection.StartLine)
                .ThenBy(d => d.Selection.StartColumn);

            var declarationsInSelection = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(d => (Selection.Selection.Contains(d.Selection) &&
                             d.QualifiedName.QualifiedModuleName == Selection.QualifiedName))
                .OrderBy(d => d.Selection.StartLine)
                .ThenBy(d => d.Selection.StartColumn);
            var referencesOfDeclarationsInSelection = (from dec in declarationsInSelection select dec.References).ToArray(); //debug purposes
            //var declarationsInEnclosingProcedure = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
            //    .Where(d => (d.));
            var sourceMethodParameters = ((IParameterizedDeclaration)TargetMethod).Parameters;
            var sourceMethodSelection = new QualifiedSelection(Selection.QualifiedName, 
                new Selection(TargetMethod.Context.Start.Line, TargetMethod.Context.Start.Column, TargetMethod.Context.Stop.Line, TargetMethod.Context.Stop.Column));
            var declarationsInParentMethod = _declarationFinderProvider.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(d => (sourceMethodSelection.Selection.Contains(d.Selection) &&
                             d.QualifiedName.QualifiedModuleName == sourceMethodSelection.QualifiedName))
                .OrderBy(d => d.Selection.StartLine)
                .ThenBy(d => d.Selection.StartColumn);

            var stopper = 10;
            //List of all variables accessed/referenced in the selection (don't need but part of finding ones we are interested in)

            //List of "inbound" variables. Parent procedure parameters + explicit dims which get referenced inside the selection.
            //Ideally excluding those declared

            //List of "outbound" variables. Any variables assigned a value in the selection which are then referenced after the selection

            //List of "inbound" only variables (to be passed ByVal)

            //List of "outbound" only variables (to be passed ByRef OR set as a return value)

            //List of "inbound" and "outbound" variables (to be passed ByRef)

            //List of neither "inbound" or "outbound" (need the declaration copied inside the selection OR moved if careful but can leave inspections to pick up unnecessary declarations)
            //Only applies if the list of inbound variables excludes those declared but not assigned before the selection

        }
        
        public string SelectedCode { get; private set; }

        private ObservableCollection<ExtractMethodParameter> _parameters;
        public ObservableCollection<ExtractMethodParameter> Parameters
        {
            get
            {
                if (_parameters == null || !_parameters.Any())
                {
                    _parameters = new ObservableCollection<ExtractMethodParameter>();
                    foreach (var declaration in SourceVariables)
                    {
                        _parameters.Add(new ExtractMethodParameter(declaration.AsTypeNameWithoutArrayDesignator,
                            ExtractMethodParameterType.ByRefParameter,
                            declaration.IdentifierName, declaration.IsArray));
                    }
                }
                return _parameters;
            }
            set => _parameters = value;
        }

        public string PreviewCode
        {
            get
            {
                _fieldsToExtract = new List<string>();
                _parametersToExtract = new List<string>();
                _variablesToExtract = new List<string>();

                foreach (var parameter in Parameters)
                {
                    switch (parameter.ParameterType)
                    {
                        case ExtractMethodParameterType.PublicModuleField:
                        case ExtractMethodParameterType.PrivateModuleField:
                            _fieldsToExtract.Add(parameter.ToString(ExtractMethodParameterFormat.DimOrParameterDeclarationWithAccessibility));
                            break;
                        case ExtractMethodParameterType.ByRefParameter:
                        case ExtractMethodParameterType.ByValParameter:
                            _parametersToExtract.Add(parameter.ToString(ExtractMethodParameterFormat.DimOrParameterDeclaration));
                            break;
                        case ExtractMethodParameterType.PrivateLocalVariable:
                        case ExtractMethodParameterType.StaticLocalVariable:
                            _variablesToExtract.Add(parameter.ToString(ExtractMethodParameterFormat.DimOrParameterDeclarationWithAccessibility));
                            break;
                        default:
                            throw new InvalidOperationException("Invalid value for ExtractParameterNewType");
                    }
                }

                var isFunction = ReturnParameter != ExtractMethodParameter.None;

                /* 
                   string.Empty are used to create blank lines
                   as the joins will create a newline each line.
                */

                var strings = new List<string>();
                var returnType = string.Empty;
                if (isFunction)
                {
                    returnType = string.Concat(Tokens.As, " ",
                        ReturnParameter.ToString(ExtractMethodParameterFormat.ReturnDeclaration) ?? Tokens.Variant);
                }
                if (_fieldsToExtract.Any())
                {
                    strings.AddRange(_fieldsToExtract);
                    strings.Add(string.Empty);
                }
                strings.Add(
                    $@"{Tokens.Private} {(isFunction ? Tokens.Function : Tokens.Sub)} {
                            NewMethodName ?? "ExtractMethod_DefaultNewMethodName" //RefactoringsUI.ExtractMethod_DefaultNewMethodName
                        }({string.Join(", ", _parametersToExtract)}) {returnType}");
                strings.AddRange(_variablesToExtract);
                if (_variablesToExtract.Any())
                {
                    strings.Add(string.Empty);
                }
                strings.AddRange(SelectedCode.Split(new[] {Environment.NewLine}, StringSplitOptions.None));
                if (isFunction)
                {
                    strings.Add(string.Empty);
                    strings.Add($"{NewMethodName} = {ReturnParameter.Name}");
                }
                strings.Add($"{Tokens.End} {(isFunction ? Tokens.Function : Tokens.Sub)}");
                return string.Join(Environment.NewLine, strings); //Indenter.Indent(strings)); //Pass to indenter
            }
        }
    }
}
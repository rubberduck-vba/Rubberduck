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
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.ExtractMethod
{
    public class ExtractMethodModel
    {
        private List<string> _fieldsList;
        private List<string> _parametersList;
        private List<string> _variablesList;

        public IEnumerable<ParserRuleContext> SelectedContexts { get; }
        public RubberduckParserState State { get; }
        public IIndenter Indenter { get; }
        public ICodeModule CodeModule { get; }
        public QualifiedSelection Selection { get; }

        public string SourceMethodName { get; private set; }
        public IEnumerable<Declaration> SourceVariables { get; private set; }
        public string NewMethodName { get; set; }
        public ExtractMethodParameter ReturnParameter { get; set; }

        public ExtractMethodModel(RubberduckParserState state, QualifiedSelection selection,
            IEnumerable<ParserRuleContext> selectedContexts, IIndenter indenter, ICodeModule codeModule)
        {
            State = state;
            Indenter = indenter;
            CodeModule = codeModule;
            Selection = selection;
            SelectedContexts = selectedContexts;
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
                NewMethodName = RubberduckUI.ExtractMethod_DefaultNewMethodName;
            }

            SelectedCode = string.Join(Environment.NewLine, SelectedContexts.Select(c => c.GetText()));

            SourceVariables = State.DeclarationFinder.UserDeclarations(DeclarationType.Variable)
                .Where(d => (Selection.Selection.Contains(d.Selection) &&
                             d.QualifiedName.QualifiedModuleName == Selection.QualifiedName) ||
                            d.References.Any(r =>
                                r.QualifiedModuleName.ComponentName == Selection.QualifiedName.ComponentName
                                && r.QualifiedModuleName.ComponentName ==
                                d.QualifiedName.QualifiedModuleName.ComponentName
                                && Selection.Selection.Contains(r.Selection)))
                .OrderBy(d => d.Selection.StartLine)
                .ThenBy(d => d.Selection.StartColumn);
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
                            ExtractMethodParameterType.PrivateLocalVariable,
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
                _fieldsList = new List<string>();
                _parametersList = new List<string>();
                _variablesList = new List<string>();

                foreach (var parameter in Parameters)
                {
                    switch (parameter.ParameterType)
                    {
                        case ExtractMethodParameterType.PublicModuleField:
                            _fieldsList.Add(parameter.ToString(ExtractMethodParameterFormat.DimOrParameterDeclarationWithAccessibility));
                            break;
                        case ExtractMethodParameterType.PrivateModuleField:
                            _fieldsList.Add(parameter.ToString(ExtractMethodParameterFormat.DimOrParameterDeclarationWithAccessibility));
                            break;
                        case ExtractMethodParameterType.ByRefParameter:
                            _parametersList.Add(parameter.ToString(ExtractMethodParameterFormat.DimOrParameterDeclaration));
                            break;
                        case ExtractMethodParameterType.ByValParameter:
                            _parametersList.Add(parameter.ToString(ExtractMethodParameterFormat.DimOrParameterDeclarationWithAccessibility));
                            break;
                        case ExtractMethodParameterType.PrivateLocalVariable:
                            _variablesList.Add(parameter.ToString(ExtractMethodParameterFormat.DimOrParameterDeclarationWithAccessibility));
                            break;
                        case ExtractMethodParameterType.StaticLocalVariable:
                            _variablesList.Add(parameter.ToString(ExtractMethodParameterFormat.DimOrParameterDeclarationWithAccessibility));
                            break;
                        default:
                            throw new InvalidOperationException("Invalid value for ExtractParameterNewType");
                    }
                }

                var isFunction = ReturnParameter != null &&
                                 !(ReturnParameter.TypeName == string.Empty &&
                                   ReturnParameter.Name == RubberduckUI.ExtractMethod_NoneSelected);

                /* 
                   string.Empty are used to create blank lines
                   as the joins will create a newline each line.
                */

                var strings = new List<string>();
                if (_fieldsList.Any())
                {
                    strings.AddRange(_fieldsList);
                    strings.Add(string.Empty);
                }
                strings.Add(
                    $@"{Tokens.Private} {(isFunction ? Tokens.Function : Tokens.Sub)} {
                            NewMethodName ?? RubberduckUI.ExtractMethod_DefaultNewMethodName
                        }({string.Join(", ", _parametersList)}) {
                            (isFunction
                                ? string.Concat(Tokens.As, " ", ReturnParameter.ToString(ExtractMethodParameterFormat.ReturnDeclaration) ?? Tokens.Variant)
                                : string.Empty)
                        }");
                strings.AddRange(_variablesList);
                if (_variablesList.Any())
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
                return string.Join(Environment.NewLine, Indenter.Indent(strings));
            }
        }
    }
}
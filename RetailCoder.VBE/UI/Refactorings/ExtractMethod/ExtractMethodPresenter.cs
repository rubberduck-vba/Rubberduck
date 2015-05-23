using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.UI.Refactorings.ExtractMethod
{
    public class ExtractMethodPresenter
    {
        private readonly IExtractMethodDialog _view;

        private readonly IEnumerable<ExtractedParameter> _input;
        private readonly IEnumerable<ExtractedParameter> _output;
        private readonly List<Declaration> _locals;
        private readonly List<Declaration> _toRemoveFromSource;

        private readonly string _selectedCode;
        private readonly QualifiedSelection _selection;

        private readonly IActiveCodePaneEditor _editor;
        private readonly Declaration _member;

        private readonly HashSet<Declaration> _usedInSelection;
        private readonly HashSet<Declaration> _usedBeforeSelection;
        private readonly HashSet<Declaration> _usedAfterSelection;

        public ExtractMethodPresenter(IActiveCodePaneEditor editor, IExtractMethodDialog view, Declaration member, QualifiedSelection selection, Declarations declarations)
        {
            _editor = editor;
            _view = view;
            _member = member;
            _selection = selection;

            _selectedCode = _editor.GetLines(selection.Selection);

            var inScopeDeclarations = declarations.Items.Where(item => item.ParentScope == member.Scope).ToList();

            var inSelection = inScopeDeclarations.SelectMany(item => item.References)
                                                 .Where(item => selection.Selection.Contains(item.Selection))
                                                 .ToList();

            _usedInSelection = new HashSet<Declaration>(inScopeDeclarations.Where(item =>
                item.References.Any(reference => inSelection.Contains(reference))));

            _usedBeforeSelection = new HashSet<Declaration>(inScopeDeclarations.Where(item => 
                item.References.Any(reference => reference.Selection.StartLine < selection.Selection.StartLine)));

            _usedAfterSelection = new HashSet<Declaration>(inScopeDeclarations.Where(item =>
                item.References.Any(reference => reference.Selection.StartLine > selection.Selection.EndLine)));

            // identifiers used inside selection and before selection (or if it's a parameter) are candidates for parameters:
            var input = inScopeDeclarations.Where(item => 
                _usedInSelection.Contains(item) && (_usedBeforeSelection.Contains(item) || item.DeclarationType == DeclarationType.Parameter)).ToList();

            // identifiers used inside selection and after selection are candidates for return values:
            var output = inScopeDeclarations.Where(item => 
                _usedInSelection.Contains(item) && _usedAfterSelection.Contains(item))
                .ToList();

            // identifiers used only inside and/or after selection are candidates for locals:
            _locals = inScopeDeclarations.Where(item => item.DeclarationType != DeclarationType.Parameter && (
                item.References.All(reference => inSelection.Contains(reference))
                || (_usedAfterSelection.Contains(item) && (!_usedBeforeSelection.Contains(item)))))
                .ToList();

            // locals that are only used in selection are candidates for being moved into the new method:
            _toRemoveFromSource = _locals.Where(item => !_usedAfterSelection.Contains(item)).ToList();

            _output = output.Select(declaration =>
                new ExtractedParameter(declaration.AsTypeName, ExtractedParameter.PassedBy.ByRef, declaration.IdentifierName));

            _input = input.Where(declaration => !output.Contains(declaration))
                .Select(declaration =>
                    new ExtractedParameter(declaration.AsTypeName, ExtractedParameter.PassedBy.ByVal, declaration.IdentifierName));
        }

        public void Show()
        {
            PrepareView();
            if (_view.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            ExtractMethod();
        }

        private void PrepareView()
        {
            _view.MethodName = "Method1";
            _view.Inputs = _input.ToList();
            _view.Outputs = _output.ToList();
            _view.Locals =
                _locals.Select(
                    variable =>
                        new ExtractedParameter(variable.AsTypeName, ExtractedParameter.PassedBy.ByVal, variable.IdentifierName))
                    .ToList();

            var returnValues = new[] {new ExtractedParameter(string.Empty, ExtractedParameter.PassedBy.ByVal)}
                .Union(_view.Outputs)
                .Union(_view.Inputs)
                .ToList();

            _view.ReturnValues = returnValues;
            _view.ReturnValue = (_output.Any() && !_output.Skip(1).Any())
                ? _output.Single()
                : returnValues.First();

            _view.RefreshPreview += _view_RefreshPreview;
            _view.OnRefreshPreview();
        }

        private void ExtractMethod()
        {
            _editor.DeleteLines(_selection.Selection);
            _editor.InsertLines(_selection.Selection.StartLine, GetMethodCall());

            var insertionLine = _member.Context.GetSelection().EndLine - _selection.Selection.LineCount + 2;
            _editor.InsertLines(insertionLine, GetExtractedMethod());

            // assumes these are declared *before* the selection...
            var offset = 0;
            foreach (var declaration in _toRemoveFromSource.OrderBy(e => e.Selection.StartLine))
            {
                var target = new Selection(
                    declaration.Selection.StartLine - offset,
                    declaration.Selection.StartColumn,
                    declaration.Selection.EndLine - offset,
                    declaration.Selection.EndColumn);

                _editor.DeleteLines(target);
                offset += declaration.Selection.LineCount;
            }
        }

        private static readonly IEnumerable<string> ValueTypes = new[]
        {
            Tokens.Boolean,
            Tokens.Byte,
            Tokens.Currency,
            Tokens.Date,
            Tokens.Decimal,
            Tokens.Double,
            Tokens.Integer,
            Tokens.Long,
            Tokens.LongLong,
            Tokens.Single,
            Tokens.String
        };

        private void _view_RefreshPreview(object sender, EventArgs e)
        {
            var hasReturnValue = _view.ReturnValue != null && _view.ReturnValue.Name != ExtractedParameter.None;
            _view.CanSetReturnValue =
                hasReturnValue && !ValueTypes.Contains(_view.ReturnValue.TypeName);

            GeneratePreview();
        }

        private void GeneratePreview()
        {
            _view.Preview = GetExtractedMethod();
        }

        private string GetMethodCall()
        {
            string result;
            var returnValueName = _view.ReturnValue.Name;
            var argsList = string.Join(", ", _view.Parameters.Select(p => p.Name));
            if (returnValueName != ExtractedParameter.None)
            {
                var setter = _view.SetReturnValue ? Tokens.Set + ' ' : string.Empty;
                result = setter + returnValueName + " = " + _view.MethodName + '(' + argsList + ')';
            }
            else
            {
                result = _view.MethodName + ' ' + argsList;
            }

            return "    " + result; // todo: smarter indentation
        }

        private string GetExtractedMethod()
        {
            var newLine = Environment.NewLine;

            var access = _view.Accessibility.ToString();
            var keyword = Tokens.Sub;
            var asTypeClause = string.Empty;

            var isFunction = _view.ReturnValue != null && _view.ReturnValue.Name != ExtractedParameter.None;
            if (isFunction)
            {
                keyword = Tokens.Function;
                asTypeClause = Tokens.As + ' ' + _view.ReturnValue.TypeName;
            }

            var parameters = "(" + string.Join(", ", _view.Parameters) + ")";

            var result = access + ' ' + keyword + ' ' + _view.MethodName + parameters + ' ' + asTypeClause + newLine;

            var localConsts = _locals.Where(e => e.DeclarationType == DeclarationType.Constant)
                .Cast<ValuedDeclaration>()
                .Select(e => "    " + Tokens.Const + ' ' + e.IdentifierName + ' ' + Tokens.As + ' ' + e.AsTypeName + " = " + e.Value);

            var localVariables = _locals.Where(e => e.DeclarationType == DeclarationType.Variable)
                .Where(e => _view.Parameters.All(param => param.Name != e.IdentifierName))
                .Select(e => e.Context)
                .Cast<VBAParser.VariableSubStmtContext>()
                .Select(e => "    " + Tokens.Dim + ' ' + e.ambiguousIdentifier().GetText() + 
                    (e.LPAREN() == null 
                        ? string.Empty 
                        : e.LPAREN().GetText() + (e.subscripts() == null ? string.Empty : e.subscripts().GetText()) + e.RPAREN().GetText()) + ' ' + 
                        (e.asTypeClause() == null ? string.Empty : e.asTypeClause().GetText()));
            var locals = string.Join(newLine, localConsts.Union(localVariables)
                            .Where(local => !_selectedCode.Contains(local)).ToArray()) + newLine;

            result += locals + _selectedCode + newLine;

            if (isFunction)
            {
                // return value by assigning the method itself:
                var setter = _view.SetReturnValue ? Tokens.Set + ' ' : string.Empty;
                result += "    " + setter + _view.MethodName + " = " + _view.ReturnValue.Name + newLine;
            }

            result += Tokens.End + ' ' + keyword + newLine;

            return newLine + result + newLine;
        }
    }
}

using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.Refactorings.ExtractMethod
{
    /// <summary>
    /// A refactoring that extracts a method (procedure or function) 
    /// out of a selection in the active code pane and 
    /// replaces the selected code with a call to the extracted method.
    /// </summary>
    public class ExtractMethodRefactoring : IRefactoring
    {
        private readonly VBE _vbe;
        private readonly IRefactoringPresenterFactory<IExtractMethodPresenter> _factory;

        public ExtractMethodRefactoring(VBE vbe, IRefactoringPresenterFactory<IExtractMethodPresenter> factory)
        {
            _vbe = vbe;
            _factory = factory;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            if (presenter == null)
            {
                OnInvalidSelection();
                return;
            }

            var model = presenter.Show();
            if (model == null)
            {
                return;
            }

            ExtractMethod(model);
        }

        public void Refactor(QualifiedSelection target)
        {
            _vbe.ActiveCodePane.CodeModule.SetSelection(target);
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            OnInvalidSelection();
        }

        private void ExtractMethod(ExtractMethodModel model)
        {
            var selection = model.Selection.Selection;

            _vbe.ActiveCodePane.CodeModule.DeleteLines(selection);
            _vbe.ActiveCodePane.CodeModule.InsertLines(selection.StartLine, GetMethodCall(model));

            var insertionLine = model.SourceMember.Context.GetSelection().EndLine - selection.LineCount + 2;
            _vbe.ActiveCodePane.CodeModule.InsertLines(insertionLine, GetExtractedMethod(model));

            // assumes these are declared *before* the selection...
            var offset = 0;
            foreach (var declaration in model.DeclarationsToMove.OrderBy(e => e.Selection.StartLine))
            {
                var target = new Selection(
                    declaration.Selection.StartLine - offset,
                    declaration.Selection.StartColumn,
                    declaration.Selection.EndLine - offset,
                    declaration.Selection.EndColumn);

                _vbe.ActiveCodePane.CodeModule.DeleteLines(target);
                offset += declaration.Selection.LineCount;
            }
        }

        private string GetMethodCall(ExtractMethodModel model)
        {
            string result;
            var returnValueName = model.Method.ReturnValue.Name;
            var argsList = string.Join(", ", model.Method.Parameters.Select(p => p.Name));
            if (returnValueName != ExtractedParameter.None)
            {
                var setter = model.Method.SetReturnValue ? Tokens.Set + ' ' : string.Empty;
                result = setter + returnValueName + " = " + model.Method.MethodName + '(' + argsList + ')';
            }
            else
            {
                result = model.Method.MethodName + ' ' + argsList;
            }

            return "    " + result; // todo: smarter indentation
        }

        /// <summary>
        /// An event that is raised when refactoring is not possible due to an invalid selection.
        /// </summary>
        public event EventHandler InvalidSelection;
        private void OnInvalidSelection()
        {
            var handler = InvalidSelection;
            if (handler != null)
            {
                handler(this, EventArgs.Empty);
            }
        }

        public static string GetExtractedMethod(ExtractMethodModel model)
        {
            var newLine = Environment.NewLine;

            var access = model.Method.Accessibility.ToString();
            var keyword = Tokens.Sub;
            var asTypeClause = string.Empty;

            var isFunction = model.Method.ReturnValue != null && model.Method.ReturnValue.Name != ExtractedParameter.None;
            if (isFunction)
            {
                keyword = Tokens.Function;
                asTypeClause = Tokens.As + ' ' + model.Method.ReturnValue.TypeName;
            }

            var parameters = "(" + string.Join(", ", model.Method.Parameters) + ")";

            var result = access + ' ' + keyword + ' ' + model.Method.MethodName + parameters + ' ' + asTypeClause + newLine;

            var localConsts = model.Locals.Where(e => e.DeclarationType == DeclarationType.Constant)
                .Cast<ValuedDeclaration>()
                .Select(e => "    " + Tokens.Const + ' ' + e.IdentifierName + ' ' + Tokens.As + ' ' + e.AsTypeName + " = " + e.Value);

            var localVariables = model.Locals.Where(e => e.DeclarationType == DeclarationType.Variable)
                .Where(e => model.Method.Parameters.All(param => param.Name != e.IdentifierName))
                .Select(e => e.Context)
                .Cast<VBAParser.VariableSubStmtContext>()
                .Select(e => "    " + Tokens.Dim + ' ' + e.identifier().GetText() +
                    (e.LPAREN() == null
                        ? string.Empty
                        : e.LPAREN().GetText() + (e.subscripts() == null ? string.Empty : e.subscripts().GetText()) + e.RPAREN().GetText()) + ' ' +
                        (e.asTypeClause() == null ? string.Empty : e.asTypeClause().GetText()));
            var locals = string.Join(newLine, localConsts.Union(localVariables)
                            .Where(local => !model.SelectedCode.Contains(local)).ToArray()) + newLine;

            result += locals + model.SelectedCode + newLine;

            if (isFunction)
            {
                // return value by assigning the method itself:
                var setter = model.Method.SetReturnValue ? Tokens.Set + ' ' : string.Empty;
                result += "    " + setter + model.Method.MethodName + " = " + model.Method.ReturnValue.Name + newLine;
            }

            result += Tokens.End + ' ' + keyword + newLine;

            return newLine + result + newLine;
        }
    }
}
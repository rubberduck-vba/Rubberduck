using System;
using Microsoft.Vbe.Interop;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldRefactoring : IRefactoring
    {
        private readonly IRefactoringPresenterFactory<IEncapsulateFieldPresenter> _factory;
        private readonly IActiveCodePaneEditor _editor;
        private EncapsulateFieldModel _model;

        public EncapsulateFieldRefactoring(IRefactoringPresenterFactory<IEncapsulateFieldPresenter> factory, IActiveCodePaneEditor editor)
        {
            _factory = factory;
            _editor = editor;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            if (presenter == null)
            {
                return;
            }

            _model = presenter.Show();

            if (_model == null) { return; }

            AddProperty();
        }

        public void Refactor(QualifiedSelection target)
        {
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            Refactor();
        }

        private void AddProperty()
        {
            UpdateReferences();

            var module = _model.TargetDeclaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
            SetFieldToPrivate(module);

            module.InsertLines(module.CountOfDeclarationLines + 1, GetPropertyText());
        }

        private void UpdateReferences()
        {
            foreach (var reference in _model.TargetDeclaration.References)
            {
                var module = reference.QualifiedModuleName.Component.CodeModule;

                var oldLine = module.Lines[reference.Selection.StartLine, 1];
                oldLine = oldLine.Remove(reference.Selection.StartColumn - 1, reference.Selection.EndColumn - reference.Selection.StartColumn);
                var newLine = oldLine.Insert(reference.Selection.StartColumn - 1, _model.PropertyName);

                module.ReplaceLine(reference.Selection.StartLine, newLine);
            }
        }

        private void SetFieldToPrivate(CodeModule module)
        {
            if (_model.TargetDeclaration.Accessibility == Accessibility.Private)
            {
                return;
            }

            RemoveField(_model.TargetDeclaration);

            var newField = "Private " + _model.TargetDeclaration.IdentifierName + " As " +
                           _model.TargetDeclaration.AsTypeName;

            module.InsertLines(module.CountOfDeclarationLines + 1, newField);

            _editor.SetSelection(_model.TargetDeclaration.QualifiedSelection);
            for (var index = 1; index <= module.CountOfDeclarationLines; index++)
            {
                if (module.Lines[index, 1].Trim() == string.Empty)
                {
                    _editor.DeleteLines(new Selection(index, 0, index, 0));
                }
            }
        }

        private void RemoveField(Declaration target)
        {
            Selection selection;
            var declarationText = target.Context.GetText();
            var multipleDeclarations = target.HasMultipleDeclarationsInStatement();

            var variableStmtContext = target.GetVariableStmtContext();

            if (!multipleDeclarations)
            {
                declarationText = variableStmtContext.GetText();
                selection = target.GetVariableStmtContextSelection();
            }
            else
            {
                selection = new Selection(target.Context.Start.Line, target.Context.Start.Column,
                    target.Context.Stop.Line, target.Context.Stop.Column);
            }

            var oldLines = _editor.GetLines(selection);

            var newLines = oldLines.Replace(" _" + Environment.NewLine, string.Empty)
                .Remove(selection.StartColumn, declarationText.Length);

            if (multipleDeclarations)
            {
                selection = target.GetVariableStmtContextSelection();
                newLines = RemoveExtraComma(_editor.GetLines(selection).Replace(oldLines, newLines),
                    target.CountOfDeclarationsInStatement(), target.IndexOfVariableDeclarationInStatement());
            }

            newLines = newLines.Replace(" _" + Environment.NewLine, string.Empty);

            _editor.DeleteLines(selection);

            if (newLines.Trim() != string.Empty)
            {
                _editor.InsertLines(selection.StartLine, newLines);
            }
        }

        private string RemoveExtraComma(string str, int numParams, int indexRemoved)
        {
            // Example use cases for this method (fields and variables):
            // Dim fizz as Boolean, dizz as Double
            // Private fizz as Boolean, dizz as Double
            // Public fizz as Boolean, _
            //        dizz as Double
            // Private fizz as Boolean _
            //         , dizz as Double _
            //         , iizz as Integer

            // Before this method is called, the parameter to be removed has 
            // already been removed.  This means 'str' will look like:
            // Dim fizz as Boolean, 
            // Private , dizz as Double
            // Public fizz as Boolean, _
            //        
            // Private  _
            //         , dizz as Double _
            //         , iizz as Integer

            // This method is responsible for removing the redundant comma
            // and returning a string similar to:
            // Dim fizz as Boolean
            // Private dizz as Double
            // Public fizz as Boolean _
            //        
            // Private  _
            //          dizz as Double _
            //         , iizz as Integer

            var commaToRemove = numParams == indexRemoved ? indexRemoved - 1 : indexRemoved;

            return str.Remove(str.NthIndexOf(',', commaToRemove), 1);
        }

        private string GetPropertyText()
        {
            var getterText = string.Join(Environment.NewLine,
                string.Format(Environment.NewLine + "Public Property Get {0}() As {1}", _model.PropertyName,
                    _model.TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", _model.PropertyName, _model.TargetDeclaration.IdentifierName),
                "End Property" + Environment.NewLine);

            var letterText = string.Join(Environment.NewLine,
                string.Format(Environment.NewLine + "Public Property Let {0}(ByVal {1} As {2})",
                    _model.PropertyName, _model.ParameterName, _model.TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", _model.TargetDeclaration.IdentifierName, _model.ParameterName),
                "End Property" + Environment.NewLine);

            var setterText = string.Join(Environment.NewLine,
                string.Format(Environment.NewLine + "Public Property Set {0}(ByVal {1} As {2})",
                    _model.PropertyName, _model.ParameterName, _model.TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", _model.TargetDeclaration.IdentifierName, _model.ParameterName),
                "End Property" + Environment.NewLine);

            return string.Join(string.Empty,
                        getterText,
                        (_model.ImplementLetSetterType ? letterText : string.Empty),
                        (_model.ImplementSetSetterType ? setterText : string.Empty)).TrimEnd() + Environment.NewLine;
        }
    }
}

using System;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    class EncapsulateFieldRefactoring : IRefactoring
    {
        private readonly IRefactoringPresenterFactory<IEncapsulateFieldPresenter> _factory;
        private EncapsulateFieldModel _model;

        public EncapsulateFieldRefactoring(IRefactoringPresenterFactory<IEncapsulateFieldPresenter> factory)
        {
            _factory = factory;
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

        private string GetPropertyText()
        {
            var getterText = string.Join(Environment.NewLine,
                string.Format(Environment.NewLine + "Public Property Get {0}() As {1}", _model.PropertyName,
                    _model.TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", _model.PropertyName, _model.TargetDeclaration.IdentifierName),
                "End Property");

            var letterText = string.Join(Environment.NewLine,
                string.Format(Environment.NewLine + "Public Property Let {0}(ByVal {1} As {2})",
                    _model.PropertyName, _model.ParameterName, _model.TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", _model.TargetDeclaration.IdentifierName, _model.ParameterName),
                "End Property");

            var setterText = string.Join(Environment.NewLine,
                string.Format(Environment.NewLine + "Public Property Set {0}(ByVal {1} As {2})",
                    _model.PropertyName, _model.ParameterName, _model.TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", _model.TargetDeclaration.IdentifierName, _model.ParameterName),
                "End Property");

            return string.Join(Environment.NewLine,
                        getterText,
                        (_model.ImplementLetSetterType ? letterText : string.Empty),
                        (_model.ImplementSetSetterType ? setterText : string.Empty));
        }
    }
}

using System;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.EncapsulateField
{
    class EncapsulateFieldRefactoring : IRefactoring
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
            var module = _model.TargetDeclaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
            module.InsertLines(module.CountOfDeclarationLines + 1, GetPropertyText());
        }

        private string GetPropertyText()
        {
            return string.Join(Environment.NewLine,
                string.Format(Environment.NewLine + "Public Property Get {0}() As {1}", _model.PropertyName,
                    _model.TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", _model.PropertyName, _model.TargetDeclaration.IdentifierName),
                "End Property" + Environment.NewLine,
                string.Format("Public Property {0} {1}({2} {3} As {4})",
                    _model.SetterTypeIsLet ? Tokens.Let : Tokens.Set, _model.PropertyName,
                    _model.ParameterModifierIsByVal ? Tokens.ByVal : Tokens.ByRef, _model.ParameterName,
                    _model.TargetDeclaration.AsTypeName),
                string.Format("    {0} = {1}", _model.TargetDeclaration.IdentifierName, _model.ParameterName),
                "End Property");
        }
    }
}

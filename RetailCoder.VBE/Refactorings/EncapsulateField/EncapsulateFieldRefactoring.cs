using System;
using Rubberduck.Common;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Selection = Rubberduck.VBEditor.Selection;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly IIndenter _indenter;
        private readonly IRefactoringPresenterFactory<IEncapsulateFieldPresenter> _factory;
        private EncapsulateFieldModel _model;

        public EncapsulateFieldRefactoring(IVBE vbe, IIndenter indenter, IRefactoringPresenterFactory<IEncapsulateFieldPresenter> factory)
        {
            _vbe = vbe;
            _indenter = indenter;
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

            QualifiedSelection? oldSelection = null;
            if (_vbe.ActiveCodePane != null)
            {
                oldSelection = _vbe.ActiveCodePane.CodeModule.GetQualifiedSelection();
            }

            AddProperty();

            if (oldSelection.HasValue)
            {
                var module = oldSelection.Value.QualifiedName.Component.CodeModule;
                var pane = module.CodePane;
                {
                    pane.Selection = oldSelection.Value.Selection;
                }
            }

            _model.State.OnParseRequested(this);
        }

        public void Refactor(QualifiedSelection target)
        {
            var pane = _vbe.ActiveCodePane;
            {
                pane.Selection = target.Selection;
            }
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            var pane = _vbe.ActiveCodePane;
            {
                pane.Selection = target.QualifiedSelection.Selection;
            }
            Refactor();
        }

        private void AddProperty()
        {
            UpdateReferences();

            var module = _model.TargetDeclaration.QualifiedName.QualifiedModuleName.Component.CodeModule;
            SetFieldToPrivate(module);

            module.InsertLines(module.CountOfDeclarationLines + 1, Environment.NewLine + GetPropertyText());
        }

        private void UpdateReferences()
        {
            foreach (var reference in _model.TargetDeclaration.References)
            {
                var module = reference.QualifiedModuleName.Component.CodeModule;
                {
                    var oldLine = module.GetLines(reference.Selection.StartLine, 1);
                    oldLine = oldLine.Remove(reference.Selection.StartColumn - 1, reference.Selection.EndColumn - reference.Selection.StartColumn);
                    var newLine = oldLine.Insert(reference.Selection.StartColumn - 1, _model.PropertyName);

                    module.ReplaceLine(reference.Selection.StartLine, newLine);
                }
            }
        }

        private void SetFieldToPrivate(ICodeModule module)
        {
            if (_model.TargetDeclaration.Accessibility == Accessibility.Private)
            {
                return;
            }

            RemoveField(_model.TargetDeclaration);

            var newField = "Private " + _model.TargetDeclaration.IdentifierName + " As " +
                           _model.TargetDeclaration.AsTypeName;

            module.InsertLines(module.CountOfDeclarationLines + 1, newField);
            var pane = module.CodePane;
            {
                pane.Selection = _model.TargetDeclaration.QualifiedSelection.Selection;
            }

            for (var index = 1; index <= module.CountOfDeclarationLines; index++)
            {
                if (module.GetLines(index, 1).Trim() == string.Empty)
                {
                    module.DeleteLines(new Selection(index, 0, index, 0));
                }
            }
        }

        private void RemoveField(Declaration target)
        {
            var module = target.QualifiedName.QualifiedModuleName.Component.CodeModule;
            module.Remove(target);
        }

        private string GetPropertyText()
        {
            var generator = new PropertyGenerator
            {
                PropertyName = _model.PropertyName,
                AsTypeName = _model.TargetDeclaration.AsTypeName,
                BackingField = _model.TargetDeclaration.IdentifierName,
                ParameterName = _model.ParameterName,
                GenerateSetter = _model.ImplementSetSetterType,
                GenerateLetter = _model.ImplementLetSetterType
            };

            var propertyTextLines = generator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, _indenter.Indent(propertyTextLines, true));
        }
    }
}

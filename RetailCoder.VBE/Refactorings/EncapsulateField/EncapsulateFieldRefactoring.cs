using System;
using System.Linq;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

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
            if (presenter == null) { return; }

            _model = presenter.Show();
            if (_model == null) { return; }

            var target = _model.TargetDeclaration;
            var rewriter = _model.State.GetRewriter(target);
            AddProperty(rewriter);

            rewriter.Rewrite();
        }

        public void Refactor(QualifiedSelection target)
        {
            var pane = _vbe.ActiveCodePane;
            pane.Selection = target.Selection;
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            var pane = _vbe.ActiveCodePane;
            pane.Selection = target.QualifiedSelection.Selection;
            Refactor();
        }

        private void AddProperty(IModuleRewriter rewriter)
        {
            UpdateReferences();
            SetFieldToPrivate(rewriter);

            var members = _model.State.DeclarationFinder
                .Members(_model.TargetDeclaration.QualifiedName.QualifiedModuleName)
                .OrderBy(declaration => declaration.QualifiedSelection)
                .ToArray();

            var property = Environment.NewLine + Environment.NewLine + GetPropertyText() + Environment.NewLine;

            var lastMember = members.LastOrDefault(m => m.DeclarationType.HasFlag(DeclarationType.Member));
            if (lastMember == null)
            {
                rewriter.InsertAtIndex(property, 1);
            }
            else
            {
                rewriter.InsertAtIndex(property, lastMember.Context.Stop.TokenIndex);
            }
        }

        private void UpdateReferences()
        {
            foreach (var reference in _model.TargetDeclaration.References)
            {
                var module = reference.QualifiedModuleName.Component.CodeModule;
                var oldLine = module.GetLines(reference.Selection.StartLine, 1);
                oldLine = oldLine.Remove(reference.Selection.StartColumn - 1, reference.Selection.EndColumn - reference.Selection.StartColumn);
                var newLine = oldLine.Insert(reference.Selection.StartColumn - 1, _model.PropertyName);

                module.ReplaceLine(reference.Selection.StartLine, newLine);
            }
        }

        private void SetFieldToPrivate(IModuleRewriter rewriter)
        {
            var target = _model.TargetDeclaration;
            if (target.Accessibility == Accessibility.Private)
            {
                return;
            }

            var newField = "Private " + _model.TargetDeclaration.IdentifierName + " As " + _model.TargetDeclaration.AsTypeName + Environment.NewLine;

            rewriter.Replace(target, newField);
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

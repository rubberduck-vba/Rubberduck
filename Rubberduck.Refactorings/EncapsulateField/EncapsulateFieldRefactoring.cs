using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Rewriter;
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
        private readonly IRewritingManager _rewritingManager;
        private readonly IRefactoringPresenterFactory<IEncapsulateFieldPresenter> _factory;
        private EncapsulateFieldModel _model;

        public EncapsulateFieldRefactoring(IVBE vbe, IIndenter indenter, IRefactoringPresenterFactory<IEncapsulateFieldPresenter> factory, IRewritingManager rewritingManager)
        {
            _vbe = vbe;
            _indenter = indenter;
            _factory = factory;
            _rewritingManager = rewritingManager;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            if (presenter == null)
            {
                return;
            }

            _model = presenter.Show();
            if (_model == null)
            {
                return;
            }

            var rewriteSession = _rewritingManager.CheckOutCodePaneSession();
            AddProperty(rewriteSession);
            rewriteSession.TryRewrite();
        }

        public void Refactor(QualifiedSelection target)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane == null || pane.IsWrappingNullReference)
                {
                    return;
                }
                pane.Selection = target.Selection;
            }
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane == null || pane.IsWrappingNullReference)
                {
                    return;
                }

                pane.Selection = target.QualifiedSelection.Selection;
            }
            Refactor();
        }

        private void AddProperty(IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(_model.TargetDeclaration.QualifiedModuleName);

            UpdateReferences(rewriteSession);
            SetFieldToPrivate(rewriter);

            var members = _model.State.DeclarationFinder
                .Members(_model.TargetDeclaration.QualifiedName.QualifiedModuleName)
                .OrderBy(declaration => declaration.QualifiedSelection);

            var fields = members.Where(d => d.DeclarationType == DeclarationType.Variable && !d.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member)).ToList();

            var property = Environment.NewLine + Environment.NewLine + GetPropertyText();
            if (members.Any(m => m.DeclarationType.HasFlag(DeclarationType.Member)))
            {
                property += Environment.NewLine;
            }

            if (_model.TargetDeclaration.Accessibility != Accessibility.Private)
            {
                var newField = $"Private {_model.TargetDeclaration.IdentifierName} As {_model.TargetDeclaration.AsTypeName}";
                if (fields.Count > 1)
                {
                    newField = Environment.NewLine + newField;
                }

                property = newField + property;
            }

            if (_model.TargetDeclaration.Accessibility == Accessibility.Private || fields.Count > 1)
            {
                rewriter.InsertAfter(fields.Last().Context.Stop.TokenIndex, property);
            }
            else
            {
                rewriter.InsertBefore(0, property);
            }
        }

        private void UpdateReferences(IRewriteSession rewriteSession)
        {
            foreach (var reference in _model.TargetDeclaration.References)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
                rewriter.Replace(reference.Context, _model.PropertyName);
            }
        }

        private void SetFieldToPrivate(IModuleRewriter rewriter)
        {
            if (_model.TargetDeclaration.Accessibility != Accessibility.Private)
            {
                rewriter.Remove(_model.TargetDeclaration);
            }
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

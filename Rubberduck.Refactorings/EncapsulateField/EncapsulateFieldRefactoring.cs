
using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.Utility;
using Environment = System.Environment;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldRefactoring : IRefactoring
    {
        private readonly RubberduckParserState _state;
        private readonly IVBE _vbe;
        private readonly IIndenter _indenter;
        private readonly IRewritingManager _rewritingManager;
        private readonly IRefactoringPresenterFactory _factory;
        private EncapsulateFieldModel _model;

        private readonly HashSet<IModuleRewriter> _referenceRewriters = new HashSet<IModuleRewriter>();
        
        public EncapsulateFieldRefactoring(RubberduckParserState state, IVBE vbe, IIndenter indenter, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager)
        {
            _state = state;
            _vbe = vbe;
            _indenter = indenter;
            _factory = factory;
            _rewritingManager = rewritingManager;
        }

        private EncapsulateFieldModel InitializeModel()
        {
            var selection = _vbe.GetActiveSelection();

            if (!selection.HasValue)
            {
                return null;
            }

            return new EncapsulateFieldModel(_state, selection.Value);
        }

        public void Refactor()
        {
            _model = InitializeModel();
            if (_model == null)
            {
                return;
            }

            using (var container = DisposalActionContainer.Create(_factory.Create<IEncapsulateFieldPresenter, EncapsulateFieldModel>(_model), p => _factory.Release(p)))
            {
                var presenter = container.Value;
                if (presenter == null)
                {
                    return;
                }

                _model = presenter.Show();
                if (_model == null)
                {
                    return;
                }

                var target = _model.TargetDeclaration;
                var rewriteSession = _rewritingManager.CheckOutCodePaneSession();
                AddProperty(rewriteSession);
                rewriteSession.TryRewrite();
            }
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
            
            var members = _model.State.DeclarationFinder
                .Members(_model.TargetDeclaration.QualifiedName.QualifiedModuleName)
                .OrderBy(declaration => declaration.QualifiedSelection);

            var fields = members.Where(d => d.DeclarationType == DeclarationType.Variable && !d.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member)).ToList();

            var property = Environment.NewLine + Environment.NewLine + GetPropertyText();
            
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
                if (_model.TargetDeclaration.Accessibility != Accessibility.Private)
                {
                    rewriter.Remove(_model.TargetDeclaration);
                }
                rewriter.InsertAfter(fields.Last().Context.Stop.TokenIndex, property);
            }
            else
            {
                rewriter.Replace(_model.TargetDeclaration.Context.GetAncestor<VBAParser.ModuleDeclarationsElementContext>(), property);
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

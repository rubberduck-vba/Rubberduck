
using System;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.Utility;
using Environment = System.Environment;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public class EncapsulateFieldRefactoring : InteractiveRefactoringBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>
    {
        private readonly RubberduckParserState _state;
        private readonly IIndenter _indenter;
        private EncapsulateFieldModel _model;
        
        public EncapsulateFieldRefactoring(RubberduckParserState state, IIndenter indenter, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService, factory)
        {
            _state = state;
            _indenter = indenter;
        }

        private EncapsulateFieldModel InitializeModel()
        {
            var activeSelection = SelectionService.ActiveSelection();

            if (!activeSelection.HasValue)
            {
                return null;
            }

            return new EncapsulateFieldModel(_state, activeSelection.Value);
        }

        public override void Refactor()
        {
            _model = InitializeModel();
            if (_model == null)
            {
                return;
            }

            using (var container = PresenterFactory(_model))
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

                var rewriteSession = RewritingManager.CheckOutCodePaneSession();
                AddProperty(rewriteSession);
                rewriteSession.TryRewrite();
            }
        }

        public override void Refactor(QualifiedSelection target)
        {
            if (!SelectionService.TrySetActiveSelection(target))
            {
                return;
            }
            Refactor();
        }

        public override void Refactor(Declaration target)
        {
            Refactor(target.QualifiedSelection);
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


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
        
        public EncapsulateFieldRefactoring(RubberduckParserState state, IIndenter indenter, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService, factory)
        {
            _state = state;
            _indenter = indenter;
        }

        public override void Refactor(QualifiedSelection target)
        {
            Refactor(InitializeModel(target));
        }

        private EncapsulateFieldModel InitializeModel(QualifiedSelection targetSelection)
        {
            return new EncapsulateFieldModel(_state, targetSelection);
        }

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            AddProperty(rewriteSession);
            rewriteSession.TryRewrite();
        }

        public override void Refactor(Declaration target)
        {
            Refactor(InitializeModel(target));
        }

        private EncapsulateFieldModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                return null;
            }

            return InitializeModel(target.QualifiedSelection);
        }

        private void AddProperty(IRewriteSession rewriteSession)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(Model.TargetDeclaration.QualifiedModuleName);

            UpdateReferences(rewriteSession);
            
            var members = Model.State.DeclarationFinder
                .Members(Model.TargetDeclaration.QualifiedName.QualifiedModuleName)
                .OrderBy(declaration => declaration.QualifiedSelection);

            var fields = members.Where(d => d.DeclarationType == DeclarationType.Variable && !d.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member)).ToList();

            var property = Environment.NewLine + Environment.NewLine + GetPropertyText();
            
            if (Model.TargetDeclaration.Accessibility != Accessibility.Private)
            {
                var newField = $"Private {Model.TargetDeclaration.IdentifierName} As {Model.TargetDeclaration.AsTypeName}";
                if (fields.Count > 1)
                {
                    newField = Environment.NewLine + newField;
                }

                property = newField + property;
            }

            if (Model.TargetDeclaration.Accessibility == Accessibility.Private || fields.Count > 1)
            {
                if (Model.TargetDeclaration.Accessibility != Accessibility.Private)
                {
                    rewriter.Remove(Model.TargetDeclaration);
                }
                rewriter.InsertAfter(fields.Last().Context.Stop.TokenIndex, property);
            }
            else
            {
                rewriter.Replace(Model.TargetDeclaration.Context.GetAncestor<VBAParser.ModuleDeclarationsElementContext>(), property);
            }
        }

        private void UpdateReferences(IRewriteSession rewriteSession)
        {
            foreach (var reference in Model.TargetDeclaration.References)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
                rewriter.Replace(reference.Context, Model.PropertyName);
            }
        }
        
        private string GetPropertyText()
        {
            var generator = new PropertyGenerator
            {
                PropertyName = Model.PropertyName,
                AsTypeName = Model.TargetDeclaration.AsTypeName,
                BackingField = Model.TargetDeclaration.IdentifierName,
                ParameterName = Model.ParameterName,
                GenerateSetter = Model.ImplementSetSetterType,
                GenerateLetter = Model.ImplementLetSetterType
            };

            var propertyTextLines = generator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, _indenter.Indent(propertyTextLines, true));
        }
    }
}

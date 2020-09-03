using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Resources;
using Rubberduck.Refactorings.CodeBlockInsert;
using System;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Refactorings.EncapsulateField;

namespace Rubberduck.Refactorings.EncapsulateFieldInsertNewCode
{
    public class EncapsulateFieldInsertNewCodeRefactoringAction : CodeOnlyRefactoringActionBase<EncapsulateFieldInsertNewCodeModel>
    {
        private const string FourSpaces = "    ";

        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IRewritingManager _rewritingManager;
        private readonly ICodeBuilder _codeBuilder;
        private readonly ICodeOnlyRefactoringAction<CodeBlockInsertModel> _codeBlockInsertRefactoringAction;
        public EncapsulateFieldInsertNewCodeRefactoringAction(
            CodeBlockInsertRefactoringAction codeBlockInsertRefactoringAction,
            IDeclarationFinderProvider declarationFinderProvider, 
            IRewritingManager rewritingManager, 
            ICodeBuilder codeBuilder)
                : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _rewritingManager = rewritingManager;
            _codeBuilder = codeBuilder;
            _codeBlockInsertRefactoringAction = codeBlockInsertRefactoringAction;
        }

        public override void Refactor(EncapsulateFieldInsertNewCodeModel model, IRewriteSession rewriteSession)
        {
            var codeSectionStartIndex = _declarationFinderProvider.DeclarationFinder
                .Members(model.QualifiedModuleName).Where(m => m.IsMember())
                .OrderBy(c => c.Selection)
                .FirstOrDefault()?.Context.Start.TokenIndex;

            var codeBlockInsertModel = new CodeBlockInsertModel()
            {
                QualifiedModuleName = model.QualifiedModuleName,
                SelectedFieldCandidates = model.SelectedFieldCandidates,
                NewContent = model.NewContent,
                CodeSectionStartIndex = codeSectionStartIndex,
                IncludeComments = model.IncludeNewContentMarker
            };

            LoadNewPropertyBlocks(codeBlockInsertModel, _codeBuilder, rewriteSession);

            _codeBlockInsertRefactoringAction.Refactor(codeBlockInsertModel, rewriteSession);
        }

        public void LoadNewPropertyBlocks(CodeBlockInsertModel model, ICodeBuilder codeBuilder, IRewriteSession rewriteSession)
        {
            if (model.IncludeComments)
            {
                model.AddContentBlock(NewContentType.PostContentMessage, RubberduckUI.EncapsulateField_PreviewMarker);
            }

            foreach (var propertyAttributes in model.SelectedFieldCandidates.SelectMany(f => f.PropertyAttributeSets))
            {
                Debug.Assert(propertyAttributes.Declaration.DeclarationType.HasFlag(DeclarationType.Variable) || propertyAttributes.Declaration.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember));

                LoadPropertyGetCodeBlock(model, propertyAttributes, codeBuilder);

                if (propertyAttributes.GenerateLetter)
                {
                    LoadPropertyLetCodeBlock(model, propertyAttributes, codeBuilder);
                }

                if (propertyAttributes.GenerateSetter)
                {
                    LoadPropertySetCodeBlock(model, propertyAttributes, codeBuilder);
                }
            }
        }

        private static void LoadPropertyLetCodeBlock(CodeBlockInsertModel model, PropertyAttributeSet propertyAttributes, ICodeBuilder codeBuilder)
        {
            var letterContent = $"{FourSpaces}{propertyAttributes.BackingField} = {propertyAttributes.ParameterName}";
            if (!codeBuilder.TryBuildPropertyLetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out var propertyLet, content: letterContent))
            {
                throw new ArgumentException();
            }
            model.AddContentBlock(NewContentType.CodeSectionBlock, propertyLet);
        }

        private static void LoadPropertySetCodeBlock(CodeBlockInsertModel model, PropertyAttributeSet propertyAttributes, ICodeBuilder codeBuilder)
        {
            var setterContent = $"{FourSpaces}{Tokens.Set} {propertyAttributes.BackingField} = {propertyAttributes.ParameterName}";
            if (!codeBuilder.TryBuildPropertySetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out var propertySet, content: setterContent))
            {
                throw new ArgumentException();
            }
            model.AddContentBlock(NewContentType.CodeSectionBlock, propertySet);
        }

        private static void LoadPropertyGetCodeBlock(CodeBlockInsertModel model, PropertyAttributeSet propertyAttributes, ICodeBuilder codeBuilder)
        {
            var getterContent = $"{propertyAttributes.PropertyName} = {propertyAttributes.BackingField}";
            if (propertyAttributes.UsesSetAssignment)
            {
                getterContent = $"{Tokens.Set} {getterContent}";
            }

            if (propertyAttributes.AsTypeName.Equals(Tokens.Variant) && !propertyAttributes.Declaration.IsArray)
            {
                getterContent = string.Join(Environment.NewLine,
                    $"{Tokens.If} IsObject({propertyAttributes.BackingField}) {Tokens.Then}",
                    $"{FourSpaces}{Tokens.Set} {propertyAttributes.PropertyName} = {propertyAttributes.BackingField}",
                    Tokens.Else,
                    $"{FourSpaces}{propertyAttributes.PropertyName} = {propertyAttributes.BackingField}",
                    $"{Tokens.End} {Tokens.If}",
                    Environment.NewLine);
            }

            if (!codeBuilder.TryBuildPropertyGetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out var propertyGet, content: $"{FourSpaces}{getterContent}"))
            {
                throw new ArgumentException();
            }

            model.AddContentBlock(NewContentType.CodeSectionBlock, propertyGet);
        }
    }
}

using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Resources;
using System;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Refactorings.EncapsulateField;
using System.Collections.Generic;

namespace Rubberduck.Refactorings.EncapsulateFieldInsertNewCode
{
    public class EncapsulateFieldInsertNewCodeRefactoringAction : CodeOnlyRefactoringActionBase<EncapsulateFieldInsertNewCodeModel>
    {
        private readonly static string _doubleSpace = $"{Environment.NewLine}{Environment.NewLine}";
        private int? _codeSectionStartIndex;
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IEncapsulateFieldCodeBuilderFactory _encapsulateFieldCodeBuilderFactory;
        public EncapsulateFieldInsertNewCodeRefactoringAction(
            IDeclarationFinderProvider declarationFinderProvider, 
            IRewritingManager rewritingManager,
            IEncapsulateFieldCodeBuilderFactory encapsulateFieldCodeBuilderFactory)
                : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _encapsulateFieldCodeBuilderFactory = encapsulateFieldCodeBuilderFactory;
        }

        public override void Refactor(EncapsulateFieldInsertNewCodeModel model, IRewriteSession rewriteSession)
        {
            _codeSectionStartIndex = _declarationFinderProvider.DeclarationFinder
                .Members(model.QualifiedModuleName).Where(m => m.IsMember())
                .OrderBy(c => c.Selection)
                .FirstOrDefault()?.Context.Start.TokenIndex;

            LoadNewPropertyBlocks(model, rewriteSession);

            InsertBlocks(model, rewriteSession);

            model.NewContentAggregator = null;
        }

        public void LoadNewPropertyBlocks(EncapsulateFieldInsertNewCodeModel model, IRewriteSession rewriteSession)
        {
            var builder = _encapsulateFieldCodeBuilderFactory.Create();
            foreach (var propertyAttributes in model.SelectedFieldCandidates.SelectMany(f => f.PropertyAttributeSets))
            {
                Debug.Assert(propertyAttributes.Declaration.DeclarationType.HasFlag(DeclarationType.Variable) || propertyAttributes.Declaration.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember));

                var (Get, Let, Set) = builder.BuildPropertyBlocks(propertyAttributes);

                var blocks = new List<string>() { Get, Let, Set };
                blocks.ForEach(s => model.NewContentAggregator.AddNewContent(NewContentType.CodeSectionBlock, s));
            }
        }

        private void InsertBlocks(EncapsulateFieldInsertNewCodeModel model, IRewriteSession rewriteSession)
        {
            var newDeclarationSectionBlock = model.NewContentAggregator.RetrieveBlock(NewContentType.UserDefinedTypeDeclaration, NewContentType.DeclarationBlock, NewContentType.CodeSectionBlock);
            if (string.IsNullOrEmpty(newDeclarationSectionBlock))
            {
                return;
            }

            var allNewContent = string.Join(_doubleSpace, new string[] { newDeclarationSectionBlock });

            var previewMarker = model.NewContentAggregator.RetrieveBlock(RubberduckUI.EncapsulateField_PreviewMarker);
            if (!string.IsNullOrEmpty(previewMarker))
            {
                allNewContent = $"{allNewContent}{Environment.NewLine}{previewMarker}";
            }

            var rewriter = rewriteSession.CheckOutModuleRewriter(model.QualifiedModuleName);

            InsertBlock(allNewContent, _codeSectionStartIndex, rewriter);
        }

        private static void InsertBlock(string content, int? insertionIndex, IModuleRewriter rewriter)
        {
            if (string.IsNullOrEmpty(content))
            {
                return;
            }

            if (insertionIndex.HasValue)
            {
                rewriter.InsertBefore(insertionIndex.Value, $"{content}{_doubleSpace}");
                return;
            }
            rewriter.InsertBefore(rewriter.TokenStream.Size - 1, $"{_doubleSpace}{content}");
        }
    }
}

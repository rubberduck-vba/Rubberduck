using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateStrategy
    {
        IEncapsulateFieldRewriteSession RefactorRewrite(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession, bool asPreview);
    }

    public abstract class EncapsulateFieldStrategyBase : IEncapsulateStrategy
    {
        protected readonly IIndenter _indenter;
        protected QualifiedModuleName _targetQMN;
        private readonly int? _codeSectionStartIndex;

        protected enum NewContentTypes { TypeDeclarationBlock, DeclarationBlock, MethodBlock, PostContentMessage };
        protected Dictionary<NewContentTypes, List<string>> _newContent { set; get; }
        private static string DoubleSpace => $"{Environment.NewLine}{Environment.NewLine}";

        public EncapsulateFieldStrategyBase(IDeclarationFinderProvider declarationFinderProvider, QualifiedModuleName qmn, IIndenter indenter)
        {
            _targetQMN = qmn;
            _indenter = indenter;

            _codeSectionStartIndex = declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN).Where(m => m.IsMember())
                .OrderBy(c => c.Selection)
                .FirstOrDefault()?.Context.Start.TokenIndex ?? null;
        }

        public IEncapsulateFieldRewriteSession RefactorRewrite(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession, bool asPreview)
        {
            ModifyFields(model, refactorRewriteSession);

            ModifyReferences(model, refactorRewriteSession);

            InsertNewContent(model, refactorRewriteSession, asPreview);

            return refactorRewriteSession;
        }

        protected abstract void ModifyFields(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession rewriteSession);

        protected abstract void ModifyReferences(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession);

        protected abstract void LoadNewDeclarationBlocks(EncapsulateFieldModel model);

        protected void RewriteReferences(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession)
        {
            foreach (var rewriteReplacement in model.SelectedFieldCandidates.SelectMany(field => field.ReferenceReplacements))
            {
                (ParserRuleContext Context, string Text) = rewriteReplacement.Value;
                var rewriter = refactorRewriteSession.CheckOutModuleRewriter(rewriteReplacement.Key.QualifiedModuleName);
                rewriter.Replace(Context, Text);
            }
        }

        protected void AddContentBlock(NewContentTypes contentType, string block)
            => _newContent[contentType].Add(block);

        private void InsertNewContent(EncapsulateFieldModel model, IEncapsulateFieldRewriteSession refactorRewriteSession, bool isPreview = false)
        {
            _newContent = new Dictionary<NewContentTypes, List<string>>
            {
                { NewContentTypes.PostContentMessage, new List<string>() },
                { NewContentTypes.DeclarationBlock, new List<string>() },
                { NewContentTypes.MethodBlock, new List<string>() },
                { NewContentTypes.TypeDeclarationBlock, new List<string>() }
            };

            LoadNewDeclarationBlocks(model);

            LoadNewPropertyBlocks(model);

            if (isPreview)
            {
                AddContentBlock(NewContentTypes.PostContentMessage, EncapsulateFieldResources.PreviewEndOfChangesMarker);
            }

            var newContentBlock = string.Join(DoubleSpace,
                            (_newContent[NewContentTypes.TypeDeclarationBlock])
                            .Concat(_newContent[NewContentTypes.DeclarationBlock])
                            .Concat(_newContent[NewContentTypes.MethodBlock])
                            .Concat(_newContent[NewContentTypes.PostContentMessage]))
                            .Trim();

            var rewriter = refactorRewriteSession.CheckOutModuleRewriter(_targetQMN);
            if (_codeSectionStartIndex.HasValue)
            {
                rewriter.InsertBefore(_codeSectionStartIndex.Value, $"{newContentBlock}{DoubleSpace}");
            }
            else
            {
                rewriter.InsertAtEndOfFile($"{DoubleSpace}{newContentBlock}");
            }
        }

        private void LoadNewPropertyBlocks(EncapsulateFieldModel model)
        {
            var propertyGenerationSpecs = model.SelectedFieldCandidates
                                                .SelectMany(f => f.PropertyAttributeSets);

            var generator = new PropertyGenerator();
            foreach (var spec in propertyGenerationSpecs)
            {
                AddContentBlock(NewContentTypes.MethodBlock, generator.AsPropertyBlock(spec, _indenter));
            }
        }
    }
}

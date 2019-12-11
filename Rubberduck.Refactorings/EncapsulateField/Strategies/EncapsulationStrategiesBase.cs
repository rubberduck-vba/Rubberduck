using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField.Strategies
{
    public interface IEncapsulateFieldStrategy
    {
        IExecutableRewriteSession GeneratePreview(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession);
        IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession);
     }

    public abstract class EncapsulateFieldStrategiesBase : IEncapsulateFieldStrategy
    {
        protected enum NewContentTypes { TypeDeclarationBlock, DeclarationBlock, MethodBlock, PostContentMessage };

        private IEncapsulateFieldValidator _validator;
        private static string DoubleSpace => $"{Environment.NewLine}{Environment.NewLine}";
        private Dictionary<NewContentTypes, List<string>> _newContent { set; get; }

        public EncapsulateFieldStrategiesBase(QualifiedModuleName qmn, IIndenter indenter, IEncapsulateFieldValidator validator)
        {
            TargetQMN = qmn;
            Indenter = indenter;
            _validator = validator;

            _newContent = new Dictionary<NewContentTypes, List<string>>
            {
                { NewContentTypes.PostContentMessage, new List<string>() },
                { NewContentTypes.DeclarationBlock, new List<string>() },
                { NewContentTypes.MethodBlock, new List<string>() },
                { NewContentTypes.TypeDeclarationBlock, new List<string>() }
            };
        }

        protected void AddCodeBlock(NewContentTypes contentType, string block)
            => _newContent[contentType].Add(block);

        protected QualifiedModuleName TargetQMN {private set; get;}

        protected IIndenter Indenter { private set; get; }

        public IExecutableRewriteSession GeneratePreview(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            if (!model.SelectedFieldCandidates.Any()) { return rewriteSession; }

            return RefactorRewrite(model, rewriteSession, asPreview: true);
        }

        public IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            if (!model.SelectedFieldCandidates.Any()) { return rewriteSession; }

            return RefactorRewrite(model, rewriteSession, asPreview: false);
        }

        protected abstract void ModifyFields(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession);

        protected abstract void LoadNewDeclarationBlocks(EncapsulateFieldModel model);

        protected virtual IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool asPreview)
        {
            ConfigureSelectedEncapsulationObjects(model);

            ModifyFields(model, rewriteSession);

            ModifyReferences(model, rewriteSession);

            RewriterRemoveWorkAround.RemoveFieldsDeclaredInLists(rewriteSession, TargetQMN);

            InsertNewContent(model, rewriteSession, asPreview);

            return rewriteSession;
        }

        protected void ConfigureSelectedEncapsulationObjects(EncapsulateFieldModel model)
        {
            foreach (var udtField in model.SelectedUDTFieldCandidates)
            {
                udtField.FieldQualifyMemberPropertyNames = model.SelectedUDTFieldCandidates.Where(f => f.AsTypeName.Equals(udtField.AsTypeName)).Count() > 1;
            }

            StageReferenceReplacementExpressions(model);
        }

        protected void ModifyReferences(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            foreach (var rewriteReplacement in model.SelectedFieldCandidates.SelectMany(fld => fld.ReferenceReplacements))
            {
                    var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, rewriteReplacement.Key.QualifiedModuleName);
                    rewriter.Replace(rewriteReplacement.Value);
            }
        }

        protected void StageReferenceReplacementExpressions(EncapsulateFieldModel model)
        {   
            foreach (var field in model.SelectedFieldCandidates)
            {
                field.LoadReferenceExpressionChanges();
            }
        }

        protected void InsertNewContent(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool postPendPreviewMessage = false)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

            LoadNewDeclarationBlocks(model);

            LoadNewPropertyBlocks(model);

            if (postPendPreviewMessage)
            {
                _newContent[NewContentTypes.PostContentMessage].Add("'<===== All Changes above this line =====>");
            }

            var newContentBlock = string.Join(DoubleSpace,
                            (_newContent[NewContentTypes.TypeDeclarationBlock])
                            .Concat(_newContent[NewContentTypes.DeclarationBlock])
                            .Concat(_newContent[NewContentTypes.MethodBlock])
                            .Concat(_newContent[NewContentTypes.PostContentMessage]))
                        .Trim();


            if (model.CodeSectionStartIndex.HasValue)
            {
                rewriter.InsertBefore(model.CodeSectionStartIndex.Value, $"{newContentBlock}{DoubleSpace}");
            }
            else
            {
                rewriter.InsertAtEndOfFile($"{DoubleSpace}{newContentBlock}");
            }
        }

        private void LoadNewPropertyBlocks(EncapsulateFieldModel model)
        {
            var propertyGenerationSpecs = model.SelectedFieldCandidates
                                                .SelectMany(f => f.PropertyGenerationSpecs);

            var generator = new PropertyGenerator();
            foreach (var spec in propertyGenerationSpecs)
            {
                AddCodeBlock(NewContentTypes.MethodBlock, generator.AsPropertyBlock(spec, Indenter));
            }
        }
    }
}

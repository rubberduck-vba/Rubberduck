using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
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
        private IEncapsulateFieldValidator _validator;
        protected static string DoubleSpace => $"{Environment.NewLine}{Environment.NewLine}";

        public EncapsulateFieldStrategiesBase(QualifiedModuleName qmn, IIndenter indenter, IEncapsulateFieldValidator validator)
        {
            TargetQMN = qmn;
            Indenter = indenter;
            _validator = validator;
        }

        protected QualifiedModuleName TargetQMN {private set; get;}

        protected IIndenter Indenter { private set; get; }

        public IExecutableRewriteSession GeneratePreview(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            if (!model.FlaggedEncapsulationFields.Any()) { return rewriteSession; }

            return RefactorRewrite(model, rewriteSession, asPreview: true);
        }

        public IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            if (!model.FlaggedEncapsulationFields.Any()) { return rewriteSession; }

            return RefactorRewrite(model, rewriteSession, asPreview: false);
        }

        protected abstract void ModifyField(IEncapsulateFieldCandidate target, IRewriteSession rewriteSession);

        protected abstract EncapsulateFieldNewContent LoadNewDeclarationBlocks(EncapsulateFieldNewContent newContent, EncapsulateFieldModel model);

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
            foreach (var udtField in model.FlaggedUDTFieldCandidates)
            {
                udtField.FieldQualifyMemberPropertyNames = model.FlaggedUDTFieldCandidates.Where(f => f.AsTypeName.Equals(udtField.AsTypeName)).Count() > 1;
            }

            StageReferenceReplacementExpressions(model);
        }

        protected void ModifyFields(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            foreach (var field in model.FlaggedEncapsulationFields)
            {
                ModifyField(field, rewriteSession);
            }
        }

        protected void ModifyReferences(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            foreach (var rewriteReplacement in model.FlaggedEncapsulationFields.SelectMany(fld => fld.ReferenceReplacements))
            {
                    var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, rewriteReplacement.Key.QualifiedModuleName);
                    rewriter.Replace(rewriteReplacement.Value);
            }
        }

        protected void StageReferenceReplacementExpressions(EncapsulateFieldModel model)
        {   
            foreach (var field in model.FlaggedEncapsulationFields)
            {
                field.LoadReferenceExpressionChanges();
            }
        }

        protected void InsertNewContent(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool postPendPreviewMessage = false)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

            var newContent = new EncapsulateFieldNewContent()
            {
                PostPendComment = postPendPreviewMessage ? "'<===== All Changes above this line =====>" : string.Empty,
            };

            newContent = LoadNewDeclarationBlocks(newContent, model);

            newContent = LoadNewPropertyBlocks(newContent, model);

            rewriter.InsertNewContent(model.CodeSectionStartIndex, newContent);
        }

        private EncapsulateFieldNewContent LoadNewPropertyBlocks(EncapsulateFieldNewContent newContent, EncapsulateFieldModel model) //, string postScript = null)
        {
            if (!model.FlaggedEncapsulationFields.Any()) { return newContent; }

            var propertyBlocks = new List<string>();
            var propertyGenerationSpecs = model.FlaggedEncapsulationFields
                                                .SelectMany(f => f.PropertyGenerationSpecs);

            foreach (var spec in propertyGenerationSpecs)
            {
                newContent.AddCodeBlock(new PropertyGenerator(spec).AsPropertyBlock(Indenter));
            }
            return newContent;
        }
    }
}

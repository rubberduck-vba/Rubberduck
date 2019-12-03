using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.EncapsulateField.Strategies
{
    public interface IEncapsulateFieldStrategy
    {
        IExecutableRewriteSession GeneratePreview(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession);
        IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession);
     }

    public abstract class EncapsulateFieldStrategiesBase : IEncapsulateFieldStrategy
    {
        private IEncapsulateFieldNamesValidator _validator;

        public EncapsulateFieldStrategiesBase(QualifiedModuleName qmn, IIndenter indenter, IEncapsulateFieldNamesValidator validator)
        {
            TargetQMN = qmn;
            Indenter = indenter;
            _validator = validator;
        }

        protected QualifiedModuleName TargetQMN {private set; get;}

        protected IIndenter Indenter { private set; get; }

        public IExecutableRewriteSession GeneratePreview(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            return RefactorRewrite(model, rewriteSession, true);
        }

        public IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            return RefactorRewrite(model, rewriteSession, false);
        }

        protected abstract void ModifyEncapsulatedVariable(IEncapsulateFieldCandidate target, IFieldEncapsulationAttributes attributes, IRewriteSession rewriteSession);

        protected abstract EncapsulateFieldNewContent LoadNewDeclarationsContent(EncapsulateFieldNewContent newContent, IEnumerable<IEncapsulateFieldCandidate> encapsulationCandidates);

        protected virtual IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool asPreview)
        {
            var nonUdtMemberFields = model.FlaggedEncapsulationFields
                    .Where(encFld => !encFld.IsUDTMember);

            foreach (var nonUdtMemberField in nonUdtMemberFields)
            {
                var attributes = nonUdtMemberField.EncapsulationAttributes;
                ModifyEncapsulatedVariable(nonUdtMemberField, attributes, rewriteSession);
                RenameReferences(nonUdtMemberField, attributes.PropertyName ?? nonUdtMemberField.Declaration.IdentifierName, rewriteSession);
            }

            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);
            RewriterRemoveWorkAround.RemoveDeclarationsFromVariableLists(rewriter);

            InsertNewContent(model.CodeSectionStartIndex, model, rewriteSession, asPreview);

            return rewriteSession;
        }

        protected void InsertNewContent(int? codeSectionStartIndex, EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool postPendPreviewMessage = false)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

            var newContent = LoadNewDeclarationsContent(new EncapsulateFieldNewContent(), model.EncapsulationFields);

            if (postPendPreviewMessage)
            {
                var postScript = "'<===== No Changes below this line =====>";
                newContent = LoadNewPropertiesContent(newContent, model.FlaggedEncapsulationFields, postScript);
            }
            else
            {
                newContent = LoadNewPropertiesContent(newContent, model.FlaggedEncapsulationFields);
            }

            rewriter.InsertNewContent(codeSectionStartIndex, newContent);

        }

        protected virtual IList<string> PropertiesContent(IEnumerable<IEncapsulateFieldCandidate> flaggedEncapsulationFields)
        {
            var textBlocks = new List<string>();
            foreach (var field in flaggedEncapsulationFields)
            {
                textBlocks.Add(BuildPropertiesTextBlock(field));
            }
            return textBlocks;
        }

        private string BuildPropertiesTextBlock(IEncapsulateFieldCandidate field)
        {
            var attributes = field.EncapsulationAttributes;
            var generator = new PropertyGenerator
            {
                PropertyName = attributes.PropertyName,
                AsTypeName = attributes.AsTypeName,
                BackingField = attributes.FieldAccessExpression(),
                ParameterName = attributes.ParameterName,
                GenerateSetter = attributes.ImplementSetSetterType,
                GenerateLetter = attributes.ImplementLetSetterType
            };

            var propertyTextLines = generator.AllPropertyCode.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
            return string.Join(Environment.NewLine, Indenter.Indent(propertyTextLines, true));
        }

        private EncapsulateFieldNewContent LoadNewPropertiesContent(EncapsulateFieldNewContent newContent, IEnumerable<IEncapsulateFieldCandidate> FlaggedEncapsulationFields, string postScript = null)
        {
            if (!FlaggedEncapsulationFields.Any()) { return newContent; }

            var theContent = string.Join($"{Environment.NewLine}{Environment.NewLine}", PropertiesContent(FlaggedEncapsulationFields));
            newContent.AddCodeBlock(theContent);
            if (postScript?.Length > 0)
            {
                newContent.AddCodeBlock($"{postScript}{Environment.NewLine}{Environment.NewLine}");
            }
            return newContent;
        }

        protected void RenameReferences(IEncapsulateFieldCandidate efd, string propertyName, IRewriteSession rewriteSession)
        {
            foreach (var reference in efd.Declaration.References)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
                rewriter.Replace(reference.Context, propertyName ?? efd.Declaration.IdentifierName);
            }
        }
    }
}

using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
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
            SetupReferenceModifications(model);

            var nonUdtMemberFields = model.FlaggedEncapsulationFields
                    .Where(encFld => !encFld.IsUDTMember);

            foreach (var nonUdtMemberField in nonUdtMemberFields)
            {
                var attributes = nonUdtMemberField.EncapsulationAttributes;
                ModifyEncapsulatedVariable(nonUdtMemberField, attributes, rewriteSession);
                RenameReferences(nonUdtMemberField, rewriteSession);
            }

            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);
            RewriterRemoveWorkAround.RemoveDeclarationsFromVariableLists(rewriter);

            InsertNewContent(model.CodeSectionStartIndex, model, rewriteSession, asPreview);

            return rewriteSession;
        }

        protected void SetupReferenceModifications(EncapsulateFieldModel model)
        {
            foreach (var field in model.FieldCandidates.Except(model.UDTFieldCandidates))
            {
                foreach (var rf in field.References)
                {
                    LoadFieldReferenceData(field, rf);
                }
            }

            foreach( var udtField in model.UDTFieldCandidates)
            {
                if (!udtField.TypeDeclarationIsPrivate)
                {
                    foreach (var rf in udtField.References)
                    {
                        LoadUDTFieldReferenceData(udtField, rf);
                    }
                }
                else
                {
                    foreach (var member in udtField.Members)
                    {
                        foreach (var rf in member.References)
                        {
                            if (rf.QualifiedModuleName == udtField.QualifiedModuleName || udtField.QualifiedModuleName.ComponentType == ComponentType.ClassModule)
                            {
                                member[rf] = udtField.EncapsulateFlag
                               ? $"{member.PropertyName}"
                               : $"{member.ReferenceExpression()}";
                            }
                            else
                            {
                                member[rf] = udtField.EncapsulateFlag
                                ? $"{udtField.QualifiedModuleName.ComponentName}.{member.PropertyName}"
                                : $"{member.ReferenceExpression()}";
                            }
                        }
                    }
                }
            }
        }

        private void LoadFieldReferenceData(IEncapsulateFieldCandidate field, IdentifierReference idRef)
        {
            if (idRef.QualifiedModuleName == field.QualifiedModuleName || field.QualifiedModuleName.ComponentType == ComponentType.ClassModule)
            {
                field[idRef] = field.ReferenceExpression();
            }
            else
            {
                if (idRef.Context.Parent is VBAParser.MemberAccessExprContext maec
                    || idRef.Context.Parent is VBAParser.WithMemberAccessExprContext wmaec)
                {
                    field[idRef] = field.ReferenceExpression();
                }
                else
                {
                    field[idRef] = $"{field.QualifiedModuleName.ComponentName}.{field.ReferenceExpression()}";
                }
            }
        }

        private void LoadUDTFieldReferenceData(IEncapsulateFieldCandidate field, IdentifierReference idRef)
        {
            if (idRef.QualifiedModuleName == field.QualifiedModuleName || field.QualifiedModuleName.ComponentType == ComponentType.ClassModule)
            {
                field[idRef] = field.ReferenceExpression();
            }
            else
            {
                if ((idRef.Context.Parent is VBAParser.MemberAccessExprContext maec
                    || idRef.Context.Parent is VBAParser.WithMemberAccessExprContext wmaec)
                    && !(idRef.Context is VBAParser.SimpleNameExprContext))
                {
                    field[idRef] = field.ReferenceExpression();
                }
                else
                {
                    field[idRef] = $"{field.QualifiedModuleName.ComponentName}.{field.ReferenceExpression()}";
                }
            }
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
                if (field is IEncapsulatedUserDefinedTypeField udtField && udtField.TypeDeclarationIsPrivate)
                {
                    foreach (var member in udtField.Members)
                    {
                        textBlocks.Add(BuildPropertiesTextBlock(member));
                    }
                    continue;
                }
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
                BackingField = attributes.PropertyAccessExpression(),
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

        protected virtual void RenameReferences(IEncapsulateFieldCandidate efd, IRewriteSession rewriteSession)
        {
            foreach (var reference in efd.Declaration.References)
            {
                if (efd.TryGetReferenceExpression(reference, out var expression))
                {
                    var replacementContext = efd.IsUDTMember
                        ? reference.Context.Parent
                        : reference.Context;

                    var rewriter = rewriteSession.CheckOutModuleRewriter(reference.QualifiedModuleName);
                    rewriter.Replace(replacementContext, expression);
                }
            }
        }
    }
}

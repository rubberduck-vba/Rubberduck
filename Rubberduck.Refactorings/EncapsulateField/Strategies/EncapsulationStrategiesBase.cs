using Rubberduck.Parsing;
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

        protected abstract void ModifyEncapsulatedField(IEncapsulateFieldCandidate target, /*IFieldEncapsulationAttributes attributes, */IRewriteSession rewriteSession);

        protected abstract EncapsulateFieldNewContent LoadNewDeclarationsContent(EncapsulateFieldNewContent newContent, IEnumerable<IEncapsulateFieldCandidate> encapsulationCandidates);

        protected virtual IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool asPreview)
        {
            var udtFieldsByTypeName = model.FlaggedUDTFieldCandidates.GroupBy((udtCandidate) => (udtCandidate as IEncapsulateFieldCandidate).AsTypeName);
            foreach (var udtField in model.FlaggedUDTFieldCandidates)
            {
                var fd = udtField as IEncapsulateFieldCandidate;
                var hasMultipleUDTFieldsOfSameType = udtFieldsByTypeName
                    .Where(group => group.Key == fd.AsTypeName).Single().Count() > 1;

                foreach (var member in udtField.Members)
                {
                    member.FieldQualifyProperty = hasMultipleUDTFieldsOfSameType;
                }
            }

            foreach (var field in model.FlaggedEncapsulationFields)
            {
                ModifyEncapsulatedField(field, rewriteSession);
            }

            ModifyReferences(model, rewriteSession);

            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);
            RewriterRemoveWorkAround.RemoveFieldsDeclaredInLists(rewriter);

            InsertNewContent(model.CodeSectionStartIndex, model, rewriteSession, asPreview);

            return rewriteSession;
        }

        private void ModifyReferences(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            SetupReferenceModifications(model);
            foreach (var field in model.FlaggedEncapsulationFields)
            {
                RenameReferences(field, rewriteSession);
                if (field is IEncapsulatedUserDefinedTypeField udtField)
                {
                    foreach (var member in udtField.Members)
                    {
                        RenameReferences(member, rewriteSession);
                    }
                }
            }
        }

        protected void SetupReferenceModifications(EncapsulateFieldModel model)
        {             
            var flaggedPrivateUDTFields = model.FlaggedUDTFieldCandidates.Where(udt => udt.TypeDeclarationIsPrivate).ToList();

            foreach (var field in model.FlaggedFieldCandidates.Except(flaggedPrivateUDTFields))
            {
                LoadFieldReferenceExpressions(field);
            }

            foreach( var udtField in flaggedPrivateUDTFields)
            {
                LoadPrivateUDTFieldReferenceExpressions(udtField);
                LoadUDTMemberReferenceExpressions(udtField);
            }
        }

        private void LoadFieldReferenceExpressions(IEncapsulateFieldCandidate field)
        {
            foreach (var idRef in field.References)
            {
                //if (idRef.QualifiedModuleName == field.QualifiedModuleName
                //    && idRef.Context.Parent is VBAParser.WithStmtContext wsc)
                //{
                //    field[idRef] = field.NewFieldName;
                //    continue;
                //}

                field[idRef] = RequiresAccessQualification(idRef)
                    ? $"{field.QualifiedModuleName.ComponentName}.{field.ReferenceExpression()}"
                    : field.ReferenceExpression();
            }
        }

        private void LoadPrivateUDTFieldReferenceExpressions(IEncapsulateFieldCandidate field)
        {
            foreach (var idRef in field.References)
            {
                if (idRef.QualifiedModuleName == field.QualifiedModuleName
                    && idRef.Context.Parent.Parent is VBAParser.WithStmtContext wsc)
                {
                    field[idRef] = field.NewFieldName;
                }
            }
        }

        private void LoadUDTMemberReferenceExpressions(IEncapsulatedUserDefinedTypeField udtField)
        {
            foreach (var member in udtField.Members)
            {
                var references = GetUDTMemberReferencesForField(member, udtField);
                foreach (var rf in references)
                {
                    var test = member.ReferenceExpression();
                    if (rf.QualifiedModuleName == udtField.QualifiedModuleName)
                    {
                        //If rf is a WithMemberAccess expression, modify the LExpr.  e.g. "With this" => "With this1"
                        if (rf.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out var wmac))
                        {
                            var wm = wmac.GetAncestor<VBAParser.WithStmtContext>();
                            var Lexpr = wm.GetChild<VBAParser.LExprContext>();
                            continue;
                        }
                        member[rf] = $"{member.PropertyName}";
                    }
                    else
                    {
                        //If rf is a WithMemberAccess expression, modify the LExpr.  e.g. "With this" => "With <qmn.ModuleName>"
                        var moduleQualifier = rf.Context.TryGetAncestor<VBAParser.WithStmtContext>(out _)
                            || rf.QualifiedModuleName == udtField.QualifiedModuleName
                            ? string.Empty
                            : $"{udtField.QualifiedModuleName.ComponentName}";

                       member[rf] = $"{moduleQualifier}.{member.PropertyName}";
                    }
                    test = member[rf];
                }
            }
        }

        private IEnumerable<IdentifierReference> GetUDTMemberReferencesForField(IEncapsulateFieldCandidate udtMember, IEncapsulatedUserDefinedTypeField field)
        {
            var refs = new List<IdentifierReference>();
            foreach (var idRef in udtMember.References)
            {
                if (idRef.Context.TryGetAncestor<VBAParser.MemberAccessExprContext>(out var mac))
                {
                    var LHS = mac.children.First();
                    switch(LHS)
                    {
                        case VBAParser.SimpleNameExprContext snec:
                            if (snec.GetText().Equals(field.IdentifierName))
                            {
                                refs.Add(idRef);
                            }
                            break;
                        case VBAParser.MemberAccessExprContext submac:
                            if (submac.children.Last() is VBAParser.UnrestrictedIdentifierContext ur &&  ur.GetText().Equals(field.IdentifierName))
                            {
                                refs.Add(idRef);
                            }
                            break;
                        case VBAParser.WithMemberAccessExprContext wmac:
                            if (wmac.children.Last().GetText().Equals(field.IdentifierName))
                            {
                                refs.Add(idRef);
                            }
                            break;
                    }
                }
                else if (idRef.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out var wmac))
                {
                    var wm = wmac.GetAncestor<VBAParser.WithStmtContext>();
                    var Lexpr = wm.GetChild<VBAParser.LExprContext>();
                    if (Lexpr.GetText().Equals(field.IdentifierName))
                    {
                        refs.Add(idRef);
                    }
                }
            }
            return refs;
        }

        private bool RequiresAccessQualification(IdentifierReference idRef)
        {
            var isLHSOfMemberAccess =
                        (idRef.Context.Parent is VBAParser.MemberAccessExprContext
                            || idRef.Context.Parent is VBAParser.WithMemberAccessExprContext)
                        && !(idRef.Context == idRef.Context.Parent.GetChild(0));// is VBAParser.SimpleNameExprContext))

            return idRef.QualifiedModuleName != idRef.Declaration.QualifiedModuleName
                        && !isLHSOfMemberAccess;
        }

        protected void InsertNewContent(int? codeSectionStartIndex, EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool postPendPreviewMessage = false)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

            var newContent = LoadNewDeclarationsContent(new EncapsulateFieldNewContent(), model.FieldCandidates);

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
                        textBlocks.Add(BuildPropertiesTextBlock(member as ISupportPropertyGenerator));
                    }
                    continue;
                }
                textBlocks.Add(BuildPropertiesTextBlock(field as ISupportPropertyGenerator));
            }
            return textBlocks;
        }

        private string BuildPropertiesTextBlock(ISupportPropertyGenerator field)
        {
            var generator = new PropertyGenerator
            {
                PropertyName = field.PropertyName,
                AsTypeName = field.AsTypeName,
                BackingField = field.PropertyAccessExpression(),
                ParameterName = field.ParameterName,
                GenerateSetter = field.ImplementSetSetterType,
                GenerateLetter = field.ImplementLetSetterType
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

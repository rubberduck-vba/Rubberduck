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
    //public interface IEncapsulateFieldStrategy
    //{
    //    IExecutableRewriteSession GeneratePreview(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession);
    //    IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession);
    //    //IEncapsulateFieldCandidate StateUDTField { set; get; }
    //}

    //public class EncapsulateFieldStrategiesBase : IEncapsulateFieldStrategy
    //{
    //    protected enum NewContentTypes { TypeDeclarationBlock, DeclarationBlock, MethodBlock, PostContentMessage };

    //    private IEncapsulateFieldValidator _validator;
    //    private static string DoubleSpace => $"{Environment.NewLine}{Environment.NewLine}";
    //    private Dictionary<NewContentTypes, List<string>> _newContent { set; get; }

    //    public EncapsulateFieldStrategiesBase(QualifiedModuleName qmn, IIndenter indenter, IEncapsulateFieldValidator validator, IEncapsulateFieldCandidate stateUDT = null)
    //    {
    //        TargetQMN = qmn;
    //        Indenter = indenter;
    //        _validator = validator;
    //        StateUDTField = stateUDT;

    //        _newContent = new Dictionary<NewContentTypes, List<string>>
    //        {
    //            { NewContentTypes.PostContentMessage, new List<string>() },
    //            { NewContentTypes.DeclarationBlock, new List<string>() },
    //            { NewContentTypes.MethodBlock, new List<string>() },
    //            { NewContentTypes.TypeDeclarationBlock, new List<string>() }
    //        };
    //    }

    //    protected void AddCodeBlock(NewContentTypes contentType, string block)
    //        => _newContent[contentType].Add(block);

    //    protected QualifiedModuleName TargetQMN {private set; get;}

    //    protected IIndenter Indenter { private set; get; }

    //    private IEncapsulateFieldCandidate StateUDTField { set; get; }

    //    public IExecutableRewriteSession GeneratePreview(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
    //    {
    //        if (!model.SelectedFieldCandidates.Any()) { return rewriteSession; }

    //        return RefactorRewrite(model, rewriteSession, asPreview: true);
    //    }

    //    public IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
    //    {
    //        if (!model.SelectedFieldCandidates.Any()) { return rewriteSession; }

    //        return RefactorRewrite(model, rewriteSession, asPreview: false);
    //    }

    //    protected virtual IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool asPreview)
    //    {
    //        ConfigureSelectedEncapsulationObjects(model);

    //        ModifyFields(model, rewriteSession);

    //        ModifyReferences(model, rewriteSession);

    //        RewriterRemoveWorkAround.RemoveFieldsDeclaredInLists(rewriteSession, TargetQMN);

    //        InsertNewContent(model, rewriteSession, asPreview);

    //        return rewriteSession;
    //    }

    //    protected void ConfigureSelectedEncapsulationObjects(EncapsulateFieldModel model)
    //    {
    //        if (model.EncapsulateWithUDT)
    //        {
    //            foreach (var field in model.SelectedFieldCandidates)
    //            {
    //                if (field is IEncapsulatedUserDefinedTypeField udt)
    //                {
    //                    udt.PropertyAccessExpression = () => $"{StateUDTField.PropertyAccessExpression()}.{udt.PropertyName}";
    //                    udt.ReferenceExpression = udt.PropertyAccessExpression;

    //                    foreach (var member in udt.Members)
    //                    {
    //                        member.PropertyAccessExpression = () => $"{udt.PropertyAccessExpression()}.{member.PropertyName}";
    //                        member.ReferenceExpression = member.PropertyAccessExpression;
    //                    }
    //                    continue;
    //                }

    //                field.PropertyAccessExpression = () => $"{StateUDTField.PropertyAccessExpression()}.{field.PropertyName}";
    //                field.ReferenceExpression = field.PropertyAccessExpression;
    //            }
    //        }

    //        foreach (var udtField in model.SelectedUDTFieldCandidates)
    //        {
    //            udtField.FieldQualifyMemberPropertyNames = model.SelectedUDTFieldCandidates.Where(f => f.AsTypeName.Equals(udtField.AsTypeName)).Count() > 1;
    //        }

    //        StageReferenceReplacementExpressions(model);
    //    }

    //    protected void ModifyReferences(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
    //    {
    //        foreach (var rewriteReplacement in model.SelectedFieldCandidates.SelectMany(fld => fld.ReferenceReplacements))
    //        {
    //                var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, rewriteReplacement.Key.QualifiedModuleName);
    //                rewriter.Replace(rewriteReplacement.Value);
    //        }
    //    }

    //    protected void StageReferenceReplacementExpressions(EncapsulateFieldModel model)
    //    {   
    //        foreach (var field in model.SelectedFieldCandidates)
    //        {
    //            field.LoadReferenceExpressionChanges();
    //        }
    //    }

    //    private void ModifyFields(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
    //    {
    //        if (model.EncapsulateWithUDT)
    //        {
    //            foreach (var field in model.SelectedFieldCandidates)
    //            {
    //                var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

    //                RewriterRemoveWorkAround.Remove(field.Declaration, rewriter);
    //                //rewriter.Remove(target.Declaration);
    //            }
    //            return;
    //        }

    //        foreach (var field in model.SelectedFieldCandidates)
    //        {
    //            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

    //            if (field.Declaration.Accessibility == Accessibility.Private && field.NewFieldName.Equals(field.Declaration.IdentifierName))
    //            {
    //                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
    //                continue;
    //            }

    //            if (field.Declaration.IsDeclaredInList())
    //            {
    //                RewriterRemoveWorkAround.Remove(field.Declaration, rewriter);
    //                //rewriter.Remove(target.Declaration);
    //                continue;
    //            }

    //            rewriter.Rename(field.Declaration, field.NewFieldName);
    //            rewriter.SetVariableVisiblity(field.Declaration, Accessibility.Private.TokenString());
    //            rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
    //        }
    //    }

    //    protected void InsertNewContent(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool postPendPreviewMessage = false)
    //    {
    //        var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, TargetQMN);

    //        LoadNewDeclarationBlocks(model);

    //        LoadNewPropertyBlocks(model);

    //        if (postPendPreviewMessage)
    //        {
    //            _newContent[NewContentTypes.PostContentMessage].Add("'<===== All Changes above this line =====>");
    //        }

    //        var newContentBlock = string.Join(DoubleSpace,
    //                        (_newContent[NewContentTypes.TypeDeclarationBlock])
    //                        .Concat(_newContent[NewContentTypes.DeclarationBlock])
    //                        .Concat(_newContent[NewContentTypes.MethodBlock])
    //                        .Concat(_newContent[NewContentTypes.PostContentMessage]))
    //                    .Trim();


    //        if (model.CodeSectionStartIndex.HasValue)
    //        {
    //            rewriter.InsertBefore(model.CodeSectionStartIndex.Value, $"{newContentBlock}{DoubleSpace}");
    //        }
    //        else
    //        {
    //            rewriter.InsertAtEndOfFile($"{DoubleSpace}{newContentBlock}");
    //        }
    //    }

    //    private void LoadNewDeclarationBlocks(EncapsulateFieldModel model)
    //    {
    //        if (model.EncapsulateWithUDT)
    //        {
    //            var udt = new UDTDeclarationGenerator(StateUDTField.AsTypeName);

    //            udt.AddMembers(model.SelectedFieldCandidates);

    //            AddCodeBlock(NewContentTypes.TypeDeclarationBlock, udt.TypeDeclarationBlock(Indenter));
    //            AddCodeBlock(NewContentTypes.DeclarationBlock, udt.FieldDeclarationBlock(StateUDTField.NewFieldName));
    //            return;
    //        }

    //        //New field declarations created here were removed from their list within ModifyFields(...)
    //        var fieldsRequiringNewDeclaration = model.SelectedFieldCandidates
    //            .Where(field => field.Declaration.IsDeclaredInList()
    //                                && field.Declaration.Accessibility != Accessibility.Private);

    //        foreach (var field in fieldsRequiringNewDeclaration)
    //        {
    //            var targetIdentifier = field.Declaration.Context.GetText().Replace(field.IdentifierName, field.NewFieldName);
    //            var newField = field.Declaration.IsTypeSpecified
    //                ? $"{Tokens.Private} {targetIdentifier}"
    //                : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {field.Declaration.AsTypeName}";

    //            AddCodeBlock(NewContentTypes.DeclarationBlock, newField);
    //        }
    //    }


    //    private void LoadNewPropertyBlocks(EncapsulateFieldModel model)
    //    {
    //        var propertyGenerationSpecs = model.SelectedFieldCandidates
    //                                            .SelectMany(f => f.PropertyGenerationSpecs);

    //        var generator = new PropertyGenerator();
    //        foreach (var spec in propertyGenerationSpecs)
    //        {
    //            AddCodeBlock(NewContentTypes.MethodBlock, generator.AsPropertyBlock(spec, Indenter));
    //        }
    //    }
    //}
}

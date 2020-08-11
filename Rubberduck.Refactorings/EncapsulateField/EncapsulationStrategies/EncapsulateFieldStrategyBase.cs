using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.Resources;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{

    public struct PropertyAttributeSet
    {
        public string PropertyName { get; set; }
        public string BackingField { get; set; }
        public string AsTypeName { get; set; }
        public string ParameterName { get; set; }
        public bool GenerateLetter { get; set; }
        public bool GenerateSetter { get; set; }
        public bool UsesSetAssignment { get; set; }
        public bool IsUDTProperty { get; set; }
        public Declaration Declaration { get; set; }
    }

    public interface IEncapsulateStrategy
    {
        IRewriteSession RefactorRewrite(IRewriteSession refactorRewriteSession, bool asPreview);
    }

    public abstract class EncapsulateFieldStrategyBase : IEncapsulateStrategy
    {
        protected readonly IIndenter _indenter;
        protected QualifiedModuleName _targetQMN;
        private readonly int? _codeSectionStartIndex;
        protected const string _defaultIndent = "    "; //4 spaces
        protected ICodeBuilder _codeBuilder;

        protected Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        protected enum NewContentTypes { TypeDeclarationBlock, DeclarationBlock, MethodBlock, PostContentMessage };
        protected Dictionary<NewContentTypes, List<string>> _newContent { set; get; }
        private static string DoubleSpace => $"{Environment.NewLine}{Environment.NewLine}";

        protected IEnumerable<IEncapsulateFieldCandidate> SelectedFields { private set; get; }

        public EncapsulateFieldStrategyBase(IDeclarationFinderProvider declarationFinderProvider, EncapsulateFieldModel model, IIndenter indenter, ICodeBuilder codeBuilder)
        {
            _targetQMN = model.QualifiedModuleName;
            _indenter = indenter;
            _codeBuilder = codeBuilder;
            SelectedFields = model.SelectedFieldCandidates.ToList();

            _codeSectionStartIndex = declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN).Where(m => m.IsMember())
                .OrderBy(c => c.Selection)
                .FirstOrDefault()?.Context.Start.TokenIndex ?? null;
        }

        public IRewriteSession RefactorRewrite(IRewriteSession refactorRewriteSession, bool asPreview)
        {
            ModifyFields(refactorRewriteSession);

            ModifyReferences(refactorRewriteSession);

            InsertNewContent(refactorRewriteSession, asPreview);

            return refactorRewriteSession;
        }

        protected abstract void ModifyFields(IRewriteSession rewriteSession);

        protected abstract void ModifyReferences(IRewriteSession refactorRewriteSession);

        protected abstract void LoadNewDeclarationBlocks();

        protected void RewriteReferences(IRewriteSession refactorRewriteSession)
        {
            foreach (var replacement in IdentifierReplacements)
            {
                (ParserRuleContext Context, string Text) = replacement.Value;
                var rewriter = refactorRewriteSession.CheckOutModuleRewriter(replacement.Key.QualifiedModuleName);
                rewriter.Replace(Context, Text);
            }
        }

        protected void AddContentBlock(NewContentTypes contentType, string block)
            => _newContent[contentType].Add(block);

        private void InsertNewContent(IRewriteSession refactorRewriteSession, bool isPreview = false)
        {
            _newContent = new Dictionary<NewContentTypes, List<string>>
            {
                { NewContentTypes.PostContentMessage, new List<string>() },
                { NewContentTypes.DeclarationBlock, new List<string>() },
                { NewContentTypes.MethodBlock, new List<string>() },
                { NewContentTypes.TypeDeclarationBlock, new List<string>() }
            };

            LoadNewDeclarationBlocks();

            LoadNewPropertyBlocks();

            if (isPreview)
            {
                AddContentBlock(NewContentTypes.PostContentMessage, RubberduckUI.EncapsulateField_PreviewMarker);
            }

            var newContentBlock = string.Join(DoubleSpace,
                            (_newContent[NewContentTypes.TypeDeclarationBlock])
                            .Concat(_newContent[NewContentTypes.DeclarationBlock])
                            .Concat(_newContent[NewContentTypes.MethodBlock])
                            .Concat(_newContent[NewContentTypes.PostContentMessage]))
                            .Trim()
                            .LimitNewlines();

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

        protected void LoadNewPropertyBlocks()
        {
            foreach (var propertyAttributes in SelectedFields.SelectMany(f => f.PropertyAttributeSets))
            {
                AddPropertyCodeBlocks(propertyAttributes);
            }
        }

        private void AddPropertyCodeBlocks(PropertyAttributeSet propertyAttributes)
        {
            Debug.Assert(propertyAttributes.Declaration.DeclarationType.HasFlag(DeclarationType.Variable) || propertyAttributes.Declaration.DeclarationType.HasFlag(DeclarationType.UserDefinedTypeMember));

            var getContent = $"{propertyAttributes.PropertyName} = {propertyAttributes.BackingField}";
            if (propertyAttributes.UsesSetAssignment)
            {
                getContent = $"{Tokens.Set} {getContent}";
            }

            if (propertyAttributes.AsTypeName.Equals(Tokens.Variant) && !propertyAttributes.Declaration.IsArray)
            {
                getContent = string.Join(Environment.NewLine,
                                    $"{Tokens.If} IsObject({propertyAttributes.BackingField}) {Tokens.Then}",
                                    $"{_defaultIndent}{Tokens.Set} {propertyAttributes.PropertyName} = {propertyAttributes.BackingField}",
                                    Tokens.Else,
                                    $"{_defaultIndent}{propertyAttributes.PropertyName} = {propertyAttributes.BackingField}",
                                    $"{Tokens.End} {Tokens.If}",
                                    Environment.NewLine);
            }

            if (!_codeBuilder.TryBuildPropertyGetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out var propertyGet, content: $"{_defaultIndent}{getContent}"))
            {
                throw new ArgumentException();
            }
            AddContentBlock(NewContentTypes.MethodBlock, propertyGet);

            if (!(propertyAttributes.GenerateLetter || propertyAttributes.GenerateSetter))
            {
                return;
            }

            if (propertyAttributes.GenerateLetter)
            {
                if (!_codeBuilder.TryBuildPropertyLetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out var propertyLet, content: $"{_defaultIndent}{propertyAttributes.BackingField} = {propertyAttributes.ParameterName}"))
                {
                    throw new ArgumentException();
                }
                AddContentBlock(NewContentTypes.MethodBlock, propertyLet);
            }

            if (propertyAttributes.GenerateSetter)
            {
                if (!_codeBuilder.TryBuildPropertySetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out var propertySet, content: $"{_defaultIndent}{Tokens.Set} {propertyAttributes.BackingField} = {propertyAttributes.ParameterName}"))
                {
                    throw new ArgumentException();
                }
                AddContentBlock(NewContentTypes.MethodBlock, propertySet);
            }
        }

        protected virtual void LoadFieldReferenceContextReplacements(IEncapsulateFieldCandidate field)
        {
            if (field is IUserDefinedTypeCandidate udt && udt.TypeDeclarationIsPrivate)
            {
                foreach (var member in udt.Members)
                {
                    foreach (var idRef in member.FieldContextReferences)
                    {
                        var replacementText = member.IdentifierForReference(idRef);
                        SetUDTMemberReferenceRewriteContent(idRef, replacementText);
                    }
                }
            }
            else
            {
                foreach (var idRef in field.Declaration.References)
                {
                    var replacementText = field.IdentifierForReference(idRef);
                    if (IsExternalReferenceRequiringModuleQualification(idRef))
                    {
                        replacementText = $"{field.QualifiedModuleName.ComponentName}.{replacementText}";
                    }
                    SetReferenceRewriteContent(idRef, replacementText);
                }
            }
        }

        protected bool IsExternalReferenceRequiringModuleQualification(IdentifierReference idRef)
        {
            var isLHSOfMemberAccess =
                        (idRef.Context.Parent is VBAParser.MemberAccessExprContext
                            || idRef.Context.Parent is VBAParser.WithMemberAccessExprContext)
                        && !(idRef.Context == idRef.Context.Parent.GetChild(0));

            return idRef.QualifiedModuleName != idRef.Declaration.QualifiedModuleName
                        && !isLHSOfMemberAccess;
        }

        protected virtual void SetReferenceRewriteContent(IdentifierReference idRef, string replacementText)
        {
            if (idRef.Context is VBAParser.IndexExprContext idxExpression)
            {
                AddIdentifierReplacement(idRef, idxExpression.children.ElementAt(0) as ParserRuleContext, replacementText);
            }
            else if (idRef.Context is VBAParser.UnrestrictedIdentifierContext
                || idRef.Context is VBAParser.SimpleNameExprContext)
            {
                AddIdentifierReplacement(idRef, idRef.Context, replacementText);
            }
            else if (idRef.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out var wmac))
            {
                AddIdentifierReplacement(idRef, wmac.GetChild<VBAParser.UnrestrictedIdentifierContext>(), replacementText);
            }
            else if (idRef.Context.TryGetAncestor<VBAParser.MemberAccessExprContext>(out var maec))
            {
                AddIdentifierReplacement(idRef, maec, replacementText);
            }
        }

        protected virtual void SetUDTMemberReferenceRewriteContent(IdentifierReference idRef, string replacementText)
        {
            if (idRef.Context is VBAParser.IndexExprContext idxExpression)
            {
                AddIdentifierReplacement(idRef, idxExpression.children.ElementAt(0) as ParserRuleContext, replacementText);
            }
            else if (idRef.Context.TryGetAncestor<VBAParser.WithMemberAccessExprContext>(out var wmac))
            {
                AddIdentifierReplacement(idRef, wmac.GetChild<VBAParser.UnrestrictedIdentifierContext>(), replacementText);
            }
            else if (idRef.Context.TryGetAncestor<VBAParser.MemberAccessExprContext>(out var maec))
            {
                AddIdentifierReplacement(idRef, maec, replacementText);
            }
        }

        private void AddIdentifierReplacement( IdentifierReference idRef, ParserRuleContext context, string replacementText)
        {
            if (IdentifierReplacements.ContainsKey(idRef))
            {
                IdentifierReplacements[idRef] = (context, replacementText);
                return;
            }
            IdentifierReplacements.Add(idRef, (context, replacementText));
        }
    }
}

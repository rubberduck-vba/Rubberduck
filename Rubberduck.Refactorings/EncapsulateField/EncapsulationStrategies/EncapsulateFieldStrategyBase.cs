using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.Resources;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
        IEncapsulateFieldRewriteSession RefactorRewrite(IEncapsulateFieldRewriteSession refactorRewriteSession, bool asPreview);
    }

    public abstract class EncapsulateFieldStrategyBase : IEncapsulateStrategy
    {
        protected readonly IIndenter _indenter;
        protected QualifiedModuleName _targetQMN;
        private readonly int? _codeSectionStartIndex;
        protected const string _defaultIndent = "    "; //4 spaces

        protected Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        protected enum NewContentTypes { TypeDeclarationBlock, DeclarationBlock, MethodBlock, PostContentMessage };
        protected Dictionary<NewContentTypes, List<string>> _newContent { set; get; }
        private static string DoubleSpace => $"{Environment.NewLine}{Environment.NewLine}";

        protected IEnumerable<IEncapsulateFieldCandidate> SelectedFields { private set; get; }

        public EncapsulateFieldStrategyBase(IDeclarationFinderProvider declarationFinderProvider, EncapsulateFieldModel model, IIndenter indenter)
        {
            _targetQMN = model.QualifiedModuleName;
            _indenter = indenter;
            SelectedFields = model.SelectedFieldCandidates;

            _codeSectionStartIndex = declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN).Where(m => m.IsMember())
                .OrderBy(c => c.Selection)
                .FirstOrDefault()?.Context.Start.TokenIndex ?? null;
        }

        public IEncapsulateFieldRewriteSession RefactorRewrite(IEncapsulateFieldRewriteSession refactorRewriteSession, bool asPreview)
        {
            ModifyFields(refactorRewriteSession);

            ModifyReferences(refactorRewriteSession);

            InsertNewContent(refactorRewriteSession, asPreview);

            return refactorRewriteSession;
        }

        protected abstract void ModifyFields(IEncapsulateFieldRewriteSession rewriteSession);

        protected abstract void ModifyReferences(IEncapsulateFieldRewriteSession refactorRewriteSession);

        protected abstract void LoadNewDeclarationBlocks();

        protected void RewriteReferences(IEncapsulateFieldRewriteSession refactorRewriteSession)
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

        private void InsertNewContent(IEncapsulateFieldRewriteSession refactorRewriteSession, bool isPreview = false)
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
                            .Trim();

            var maxConsecutiveNewLines = 3;
            var target = string.Join(string.Empty, Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines).ToList());
            var replacement = string.Join(string.Empty, Enumerable.Repeat(Environment.NewLine, maxConsecutiveNewLines - 1).ToList());
            for (var counter = 1; counter < 10 && newContentBlock.Contains(target); counter++)
            {
                newContentBlock = newContentBlock.Replace(target, replacement);
            }


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

        protected virtual void LoadNewPropertyBlocks()
        {
            foreach (var attributeSet in SelectedFields.SelectMany(f => f.PropertyAttributeSets))
            {
                if (attributeSet.Declaration is VariableDeclaration || attributeSet.Declaration.DeclarationType.Equals(DeclarationType.UserDefinedTypeMember))
                {
                    var getContent = $"{attributeSet.PropertyName} = {attributeSet.BackingField}";
                    if (attributeSet.UsesSetAssignment)
                    {
                        getContent = $"{Tokens.Set} {getContent}";
                    }
                    if (attributeSet.AsTypeName.Equals(Tokens.Variant) && !attributeSet.Declaration.IsArray)
                    {
                        getContent = string.Join(Environment.NewLine,
                                            $"{Tokens.If} IsObject({attributeSet.BackingField}) {Tokens.Then}",
                                            $"{_defaultIndent}{Tokens.Set} {attributeSet.PropertyName} = {attributeSet.BackingField}",
                                            Tokens.Else,
                                            $"{_defaultIndent}{attributeSet.PropertyName} = {attributeSet.BackingField}",
                                            $"{Tokens.End} {Tokens.If}",
                                            Environment.NewLine);
                    }

                    AddContentBlock(NewContentTypes.MethodBlock, attributeSet.Declaration.FieldToPropertyBlock(DeclarationType.PropertyGet, attributeSet.PropertyName, content: $"{_defaultIndent}{getContent}"));

                    if (attributeSet.GenerateLetter)
                    {
                        AddContentBlock(NewContentTypes.MethodBlock, attributeSet.Declaration.FieldToPropertyBlock(DeclarationType.PropertyLet, attributeSet.PropertyName, content: $"{_defaultIndent}{attributeSet.BackingField} = {attributeSet.ParameterName}"));
                    }
                    if (attributeSet.GenerateSetter)
                    {
                        AddContentBlock(NewContentTypes.MethodBlock, attributeSet.Declaration.FieldToPropertyBlock(DeclarationType.PropertySet, attributeSet.PropertyName, content: $"{_defaultIndent}{Tokens.Set} {attributeSet.BackingField} = {attributeSet.ParameterName}"));
                    }
                }
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

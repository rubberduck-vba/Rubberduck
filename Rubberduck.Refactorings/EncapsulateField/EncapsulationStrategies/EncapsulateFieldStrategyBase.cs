using Antlr4.Runtime;
using Rubberduck.Parsing;
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

        protected Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        protected enum NewContentTypes { TypeDeclarationBlock, DeclarationBlock, MethodBlock, PostContentMessage };
        protected Dictionary<NewContentTypes, List<string>> _newContent { set; get; }
        private static string DoubleSpace => $"{Environment.NewLine}{Environment.NewLine}";

        protected IEnumerable<IEncapsulatableField> SelectedFields { private set; get; }

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
            foreach (var replacement in IdentifierReplacements)
            {
                (ParserRuleContext Context, string Text) = replacement.Value;
                var rewriter = refactorRewriteSession.CheckOutModuleRewriter(replacement.Key.QualifiedModuleName);
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
                AddContentBlock(NewContentTypes.PostContentMessage, EncapsulateFieldResources.PreviewMarker);
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

        protected virtual void LoadNewPropertyBlocks(EncapsulateFieldModel model)
        {
            var propertyGenerationSpecs = SelectedFields.SelectMany(f => f.PropertyAttributeSets);

            var generator = new PropertyGenerator();
            foreach (var spec in propertyGenerationSpecs)
            {
                AddContentBlock(NewContentTypes.MethodBlock, generator.AsPropertyBlock(spec, _indenter));
            }
        }

        protected virtual void LoadFieldReferenceContextReplacements(IEncapsulatableField field)
        {
            if (field is IUserDefinedTypeCandidate udt && udt.TypeDeclarationIsPrivate)
            {
                foreach (var member in udt.Members)
                {
                    foreach (var idRef in member.ParentContextReferences)
                    {
                        var replacementText = member.ReferenceAccessor(idRef);
                        SetUDTMemberReferenceRewriteContent(idRef, replacementText);
                    }
                }
            }
            else
            {
                foreach (var idRef in field.Declaration.References)
                {
                    var replacementText = field.ReferenceAccessor(idRef);
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

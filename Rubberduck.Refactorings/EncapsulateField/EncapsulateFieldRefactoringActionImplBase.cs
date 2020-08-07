using Antlr4.Runtime;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.EncapsulateField.Extensions;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public abstract class EncapsulateFieldRefactoringActionImplBase : CodeOnlyRefactoringActionBase<EncapsulateFieldModel>
    {
        private const string _defaultIndent = "    "; //4 spaces
        private static string _doubleSpace = $"{Environment.NewLine}{Environment.NewLine}";

        protected enum NewContentType
        {
            TypeDeclarationBlock,
            DeclarationBlock,
            MethodBlock,
            PostContentMessage
        };

        protected readonly IDeclarationFinderProvider _declarationFinderProvider;
        protected readonly IIndenter _indenter;
        protected readonly ICodeBuilder _codeBuilder;

        protected QualifiedModuleName _targetQMN;
        protected int? _codeSectionStartIndex;

        protected Dictionary<NewContentType, List<string>> _newContent { set; get; }
        protected Dictionary<IdentifierReference, (ParserRuleContext, string)> IdentifierReplacements { get; } = new Dictionary<IdentifierReference, (ParserRuleContext, string)>();

        public EncapsulateFieldRefactoringActionImplBase(
                IDeclarationFinderProvider declarationFinderProvider,
                IIndenter indenter,
                IRewritingManager rewritingManager,
                ICodeBuilder codeBuilder)
            : base(rewritingManager)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _indenter = indenter;
            _codeBuilder = codeBuilder;
        }

        protected IEnumerable<IEncapsulateFieldCandidate> SelectedFields { set; get; }

        protected IRewriteSession RefactorImpl(EncapsulateFieldModel model, IRewriteSession rewriteSession)
        {
            InitializeRefactoringAction(model);

            ModifyFields(rewriteSession);

            ModifyReferences(rewriteSession);

            InsertNewContent(rewriteSession);

            return rewriteSession;
        }

        protected abstract void ModifyFields(IRewriteSession rewriteSession);

        protected abstract void ModifyReferences(IRewriteSession refactorRewriteSession);

        protected abstract void LoadNewDeclarationBlocks();

        protected void RewriteReferences(IRewriteSession rewriteSession)
        {
            foreach (var replacement in IdentifierReplacements)
            {
                (ParserRuleContext Context, string Text) = replacement.Value;
                var rewriter = rewriteSession.CheckOutModuleRewriter(replacement.Key.QualifiedModuleName);
                rewriter.Replace(Context, Text);
            }
        }

        protected void AddContentBlock(NewContentType contentType, string block)
            => _newContent[contentType].Add(block);

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

        private void InitializeRefactoringAction(EncapsulateFieldModel model)
        {
            _targetQMN = model.QualifiedModuleName;

            _codeSectionStartIndex = _declarationFinderProvider.DeclarationFinder
                .Members(model.QualifiedModuleName).Where(m => m.IsMember())
                .OrderBy(c => c.Selection)
                .FirstOrDefault()?.Context.Start.TokenIndex ?? null;

            SelectedFields = model.SelectedFieldCandidates;
        }

        private void AddIdentifierReplacement(IdentifierReference idRef, ParserRuleContext context, string replacementText)
        {
            if (IdentifierReplacements.ContainsKey(idRef))
            {
                IdentifierReplacements[idRef] = (context, replacementText);
                return;
            }
            IdentifierReplacements.Add(idRef, (context, replacementText));
        }

        private void InsertNewContent(IRewriteSession refactorRewriteSession)
        {
            _newContent = new Dictionary<NewContentType, List<string>>
            {
                { NewContentType.PostContentMessage, new List<string>() },
                { NewContentType.DeclarationBlock, new List<string>() },
                { NewContentType.MethodBlock, new List<string>() },
                { NewContentType.TypeDeclarationBlock, new List<string>() }
            };

            LoadNewDeclarationBlocks();

            LoadNewPropertyBlocks();

            var newContentBlock = string.Join(_doubleSpace,
                            (_newContent[NewContentType.TypeDeclarationBlock])
                            .Concat(_newContent[NewContentType.DeclarationBlock])
                            .Concat(_newContent[NewContentType.MethodBlock])
                            .Concat(_newContent[NewContentType.PostContentMessage]))
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
                rewriter.InsertBefore(_codeSectionStartIndex.Value, $"{newContentBlock}{_doubleSpace}");
            }
            else
            {
                rewriter.InsertAtEndOfFile($"{_doubleSpace}{newContentBlock}");
            }
        }

        protected void LoadNewPropertyBlocks()
        {
            foreach (var propertyAttributes in SelectedFields.SelectMany(f => f.PropertyAttributeSets))
            {
                AddPropertyCodeBlocks(propertyAttributes);
            }
        }
        /// <summary>
        /// RemoveFields handles the special case of field declaration removal where 
        /// each field of a VariableListStmtContext is specified for removal.  In this
        /// special case the Parent context is removed rather than the individual declarations.
        /// </summary>
        protected static void RemoveFields(IEnumerable<Declaration> toRemove, IRewriteSession rewriteSession)
        {
            if (!toRemove.Any()) { return; }

            var fieldsByListContext = toRemove.Distinct().GroupBy(f => f.Context.GetAncestor<VBAParser.VariableListStmtContext>());

            var rewriter = rewriteSession.CheckOutModuleRewriter(toRemove.First().QualifiedModuleName);
            foreach (var fieldsGroup in fieldsByListContext)
            {
                var variables = fieldsGroup.Key.children.Where(ch => ch is VBAParser.VariableSubStmtContext);
                if (variables.Count() == fieldsGroup.Count())
                {
                    rewriter.Remove(fieldsGroup.Key.Parent);
                    continue;
                }

                foreach (var target in fieldsGroup)
                {
                    rewriter.Remove(target);
                }
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
            AddContentBlock(NewContentType.MethodBlock, propertyGet);

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
                AddContentBlock(NewContentType.MethodBlock, propertyLet);
            }

            if (propertyAttributes.GenerateSetter)
            {
                if (!_codeBuilder.TryBuildPropertySetCodeBlock(propertyAttributes.Declaration, propertyAttributes.PropertyName, out var propertySet, content: $"{_defaultIndent}{Tokens.Set} {propertyAttributes.BackingField} = {propertyAttributes.ParameterName}"))
                {
                    throw new ArgumentException();
                }
                AddContentBlock(NewContentType.MethodBlock, propertySet);
            }
        }
    }
}

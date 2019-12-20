using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor.Utility;
using System.Collections.Generic;
using System;
using Rubberduck.Parsing;
using Rubberduck.Refactorings.Common;
using System.IO;
using Antlr4.Runtime;

namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldRefactoringTestAccess
    {
        EncapsulateFieldModel TestUserInteractionOnly(Declaration target, Func<EncapsulateFieldModel, EncapsulateFieldModel> userInteraction);
    }

    public class EncapsulateFieldRefactoring : InteractiveRefactoringBase<IEncapsulateFieldPresenter, EncapsulateFieldModel>, IEncapsulateFieldRefactoringTestAccess
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly ISelectedDeclarationProvider _selectedDeclarationProvider;
        private readonly IIndenter _indenter;
        private QualifiedModuleName _targetQMN;
        private EncapsulateFieldElementFactory _encapsulationCandidateFactory;

        private enum NewContentTypes { TypeDeclarationBlock, DeclarationBlock, MethodBlock, PostContentMessage };
        private Dictionary<NewContentTypes, List<string>> _newContent { set; get; }

        private int? _codeSectionStartIndex;

        private static string DoubleSpace => $"{Environment.NewLine}{Environment.NewLine}";

        public EncapsulateFieldRefactoring(
            IDeclarationFinderProvider declarationFinderProvider,
            IIndenter indenter,
            IRefactoringPresenterFactory factory,
            IRewritingManager rewritingManager,
            ISelectionProvider selectionProvider,
            ISelectedDeclarationProvider selectedDeclarationProvider,
            IUiDispatcher uiDispatcher)
        :base(rewritingManager, selectionProvider, factory, uiDispatcher)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _selectedDeclarationProvider = selectedDeclarationProvider;
            _indenter = indenter;

            _codeSectionStartIndex = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN).Where(m => m.IsMember())
                .OrderBy(c => c.Selection)
                .FirstOrDefault()?.Context.Start.TokenIndex ?? null;
        }

        public EncapsulateFieldModel Model { set; get; }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var selectedDeclaration = _selectedDeclarationProvider.SelectedDeclaration(targetSelection);
            if (selectedDeclaration == null
                || selectedDeclaration.DeclarationType != DeclarationType.Variable
                || selectedDeclaration.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
            {
                return null;
            }

            return selectedDeclaration;
        }

        public EncapsulateFieldModel TestUserInteractionOnly(Declaration target, Func<EncapsulateFieldModel, EncapsulateFieldModel> userInteraction)
        {
            var model = InitializeModel(target);
            return userInteraction(model);
        }

        protected override EncapsulateFieldModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!target.DeclarationType.Equals(DeclarationType.Variable))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            _targetQMN = target.QualifiedModuleName;

            var validator = new EncapsulateFieldNamesValidator(_declarationFinderProvider);
            _encapsulationCandidateFactory = new EncapsulateFieldElementFactory(_declarationFinderProvider, _targetQMN, validator);

            var candidates = _encapsulationCandidateFactory.CreateEncapsulationCandidates();
            var selected = candidates.Single(c => c.Declaration == target);
            selected.EncapsulateFlag = true;

            Model = new EncapsulateFieldModel(
                                target,
                                candidates,
                                _encapsulationCandidateFactory.CreateStateUDTField(),
                                PreviewRewrite);

            _codeSectionStartIndex = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN).Where(m => m.IsMember())
                .OrderBy(c => c.Selection)
                            .FirstOrDefault()?.Context.Start.TokenIndex ?? null;

            return Model;
        }

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            var rewriteSession = RefactorRewrite(model, RewritingManager.CheckOutCodePaneSession());

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private string PreviewRewrite(EncapsulateFieldModel model)
        {
            var rewriteSession = GeneratePreview(model, RewritingManager.CheckOutCodePaneSession());

            var previewRewriter = rewriteSession.CheckOutModuleRewriter(_targetQMN);

            return previewRewriter.GetText(maxConsecutiveNewLines: 3);
        }

        private IStateUDT StateUDTField
            => Model.EncapsulateWithUDT ? Model.StateUDTField : null;


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

        private IExecutableRewriteSession RefactorRewrite(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool asPreview)
        {
            _newContent = new Dictionary<NewContentTypes, List<string>>
            {
                { NewContentTypes.PostContentMessage, new List<string>() },
                { NewContentTypes.DeclarationBlock, new List<string>() },
                { NewContentTypes.MethodBlock, new List<string>() },
                { NewContentTypes.TypeDeclarationBlock, new List<string>() }
            };

            ModifyFields(model, rewriteSession);

            ModifyReferences(model, rewriteSession);

            RewriterRemoveWorkAround.RemoveFieldsDeclaredInLists(rewriteSession, _targetQMN);

            InsertNewContent(model, rewriteSession, asPreview);

            return rewriteSession;
        }

        private void ModifyReferences(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            var stateUDT = model.EncapsulateWithUDT
                ? model.StateUDTField
                : null;

            foreach (var field in model.SelectedFieldCandidates)
            {
                field.StageFieldReferenceReplacements(stateUDT);
            }

            foreach (var rewriteReplacement in model.SelectedFieldCandidates.SelectMany(field => field.ReferenceReplacements))
            {
                (ParserRuleContext Context, string Text) = rewriteReplacement.Value;
                var rewriter = rewriteSession.CheckOutModuleRewriter(rewriteReplacement.Key.QualifiedModuleName);
                rewriter.Replace(Context, Text);
            }
        }

        private void ModifyFields(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            if (model.EncapsulateWithUDT)
            {
                foreach (var field in model.SelectedFieldCandidates)
                {
                    var rewriter = rewriteSession.CheckOutModuleRewriter(_targetQMN);

                    RewriterRemoveWorkAround.Remove(field.Declaration, rewriter);
                }
                return;
            }

            foreach (var field in model.SelectedFieldCandidates)
            {
                var rewriter = rewriteSession.CheckOutModuleRewriter(_targetQMN);

                if (field.Declaration.HasPrivateAccessibility() && field.FieldIdentifier.Equals(field.Declaration.IdentifierName))
                {
                    rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
                    continue;
                }

                if (field.Declaration.IsDeclaredInList() && !field.Declaration.HasPrivateAccessibility())
                {
                    //FIXME: the call here should be: rewriter.Remove(field.Declaration);
                    RewriterRemoveWorkAround.Remove(field.Declaration, rewriter);
                    continue;
                }

                rewriter.Rename(field.Declaration, field.FieldIdentifier);
                rewriter.SetVariableVisiblity(field.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
            }
        }

        private void InsertNewContent(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool postPendPreviewMessage = false)
        {
            var rewriter = rewriteSession.CheckOutModuleRewriter(_targetQMN);

            LoadNewDeclarationBlocks(model);

            LoadNewPropertyBlocks(model);

            if (postPendPreviewMessage)
            {
                _newContent[NewContentTypes.PostContentMessage].Add(EncapsulateFieldResources.PreviewEndOfChangesMarker);
            }

            var newContentBlock = string.Join(DoubleSpace,
                            (_newContent[NewContentTypes.TypeDeclarationBlock])
                            .Concat(_newContent[NewContentTypes.DeclarationBlock])
                            .Concat(_newContent[NewContentTypes.MethodBlock])
                            .Concat(_newContent[NewContentTypes.PostContentMessage]))
                        .Trim();


            if (_codeSectionStartIndex.HasValue)
            {
                rewriter.InsertBefore(_codeSectionStartIndex.Value, $"{newContentBlock}{DoubleSpace}");
            }
            else
            {
                rewriter.InsertAtEndOfFile($"{DoubleSpace}{newContentBlock}");
            }
        }

        private void LoadNewDeclarationBlocks(EncapsulateFieldModel model)
        {
            if (model.EncapsulateWithUDT)
            {
                var stateUDT = StateUDTField as IStateUDT;
                stateUDT.AddMembers(model.SelectedFieldCandidates);

                AddCodeBlock(NewContentTypes.TypeDeclarationBlock, stateUDT.TypeDeclarationBlock(_indenter));
                AddCodeBlock(NewContentTypes.DeclarationBlock, stateUDT.FieldDeclarationBlock);
                return;
            }

            //New field declarations created here were removed from their list within ModifyFields(...)
            var fieldsRequiringNewDeclaration = model.SelectedFieldCandidates
                .Where(field => field.Declaration.IsDeclaredInList()
                                    && field.Declaration.Accessibility != Accessibility.Private);

            foreach (var field in fieldsRequiringNewDeclaration)
            {
                var targetIdentifier = field.Declaration.Context.GetText().Replace(field.IdentifierName, field.FieldIdentifier);
                var newField = field.Declaration.IsTypeSpecified
                    ? $"{Tokens.Private} {targetIdentifier}"
                    : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {field.Declaration.AsTypeName}";

                AddCodeBlock(NewContentTypes.DeclarationBlock, newField);
            }
        }

        private void LoadNewPropertyBlocks(EncapsulateFieldModel model)
        {
            var propertyGenerationSpecs = model.SelectedFieldCandidates
                                                .SelectMany(f => f.PropertyAttributeSets);

            var generator = new PropertyGenerator();
            foreach (var spec in propertyGenerationSpecs)
            {
                AddCodeBlock(NewContentTypes.MethodBlock, generator.AsPropertyBlock(spec, _indenter));
            }
        }

        private void AddCodeBlock(NewContentTypes contentType, string block)
            => _newContent[contentType].Add(block);
    }
}

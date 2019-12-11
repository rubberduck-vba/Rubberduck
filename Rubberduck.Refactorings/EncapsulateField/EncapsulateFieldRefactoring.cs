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
using Rubberduck.Refactorings.EncapsulateField.Strategies;

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
        private EncapsulationCandidateFactory _encapsulationCandidateFactory;
        private IEncapsulateFieldNamesValidator _validator;

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
            _validator = new EncapsulateFieldNamesValidator(_declarationFinderProvider);//, () => FieldCandidates);
            //_useNewValidationScheme = File.Exists("C:\\Users\\Brian\\Documents\\UseNewUDTStructure.txt");

            _codeSectionStartIndex = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN).Where(m => m.IsMember())
                .OrderBy(c => c.Selection)
                .FirstOrDefault()?.Context.Start.TokenIndex ?? null; ;

            //_codeSectionStartIndex = moduleMembers.OrderBy(c => c.Selection)
            //                            .FirstOrDefault()?.Context.Start.TokenIndex ?? null;
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

            _encapsulationCandidateFactory = new EncapsulationCandidateFactory(_declarationFinderProvider, _targetQMN, _validator);

            Model = new EncapsulateFieldModel(
                                target,
                                //_declarationFinderProvider,
                                //_indenter,
                                _encapsulationCandidateFactory.CreateEncapsulationCandidates(),
                                _encapsulationCandidateFactory.CreateStateUDTField(),
                                PreviewRewrite);

            //Model.StateUDTField = _encapsulationCandidateFactory.CreateStateUDTField();

            var moduleMembers = _declarationFinderProvider.DeclarationFinder
                .Members(_targetQMN).Where(m => m.IsMember());

            _codeSectionStartIndex
                = moduleMembers.OrderBy(c => c.Selection)
                            .FirstOrDefault()?.Context.Start.TokenIndex ?? null;

            //Model.CodeSectionStartIndex = codeSectionStartIndex;

            return Model;
        }

        protected override void RefactorImpl(EncapsulateFieldModel model)
        {
            //var strategy = new EncapsulateFieldStrategiesBase(_targetQMN, _indenter, _validator, StateUDTField);
            var rewriteSession = RefactorRewrite(model, RewritingManager.CheckOutCodePaneSession());

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
        }

        private string PreviewRewrite(EncapsulateFieldModel model)
        {
            //var strategy = new EncapsulateFieldStrategiesBase(_targetQMN, _indenter, _validator, StateUDTField);
            var scratchPadRewriteSession = GeneratePreview(model, RewritingManager.CheckOutCodePaneSession());

            var previewRewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(scratchPadRewriteSession, _targetQMN);

            return previewRewriter.GetText(maxConsecutiveNewLines: 3);
        }

        private IEncapsulateFieldCandidate StateUDTField
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

            ConfigureSelectedEncapsulationObjects(model);

            ModifyFields(model, rewriteSession);

            ModifyReferences(model, rewriteSession);

            RewriterRemoveWorkAround.RemoveFieldsDeclaredInLists(rewriteSession, _targetQMN);

            InsertNewContent(model, rewriteSession, asPreview);

            return rewriteSession;
        }

        private void ConfigureSelectedEncapsulationObjects(EncapsulateFieldModel model)
        {
            if (model.EncapsulateWithUDT)
            {
                foreach (var field in model.SelectedFieldCandidates)
                {
                    if (field is IEncapsulatedUserDefinedTypeField udt)
                    {
                        udt.PropertyAccessExpression = () => $"{StateUDTField.PropertyAccessExpression()}.{udt.PropertyName}";
                        udt.ReferenceExpression = udt.PropertyAccessExpression;

                        foreach (var member in udt.Members)
                        {
                            member.PropertyAccessExpression = () => $"{udt.PropertyAccessExpression()}.{member.PropertyName}";
                            member.ReferenceExpression = member.PropertyAccessExpression;
                        }
                        continue;
                    }

                    field.PropertyAccessExpression = () => $"{StateUDTField.PropertyAccessExpression()}.{field.PropertyName}";
                    field.ReferenceExpression = field.PropertyAccessExpression;
                }
            }

            foreach (var udtField in model.SelectedUDTFieldCandidates)
            {
                udtField.FieldQualifyMemberPropertyNames = model.SelectedUDTFieldCandidates.Where(f => f.AsTypeName.Equals(udtField.AsTypeName)).Count() > 1;
            }

            //StageReferenceReplacementExpressions(model);
            foreach (var field in model.SelectedFieldCandidates)
            {
                field.LoadReferenceExpressionChanges();
            }
        }

        private void ModifyReferences(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            foreach (var rewriteReplacement in model.SelectedFieldCandidates.SelectMany(fld => fld.ReferenceReplacements))
            {
                var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, rewriteReplacement.Key.QualifiedModuleName);
                rewriter.Replace(rewriteReplacement.Value);
            }
        }

        //private void StageReferenceReplacementExpressions(EncapsulateFieldModel model)
        //{
        //    foreach (var field in model.SelectedFieldCandidates)
        //    {
        //        field.LoadReferenceExpressionChanges();
        //    }
        //}

        private void ModifyFields(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession)
        {
            if (model.EncapsulateWithUDT)
            {
                foreach (var field in model.SelectedFieldCandidates)
                {
                    var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, _targetQMN);

                    RewriterRemoveWorkAround.Remove(field.Declaration, rewriter);
                    //rewriter.Remove(target.Declaration);
                }
                return;
            }

            foreach (var field in model.SelectedFieldCandidates)
            {
                var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, _targetQMN);

                if (field.Declaration.Accessibility == Accessibility.Private && field.NewFieldName.Equals(field.Declaration.IdentifierName))
                {
                    rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
                    continue;
                }

                if (field.Declaration.IsDeclaredInList())
                {
                    RewriterRemoveWorkAround.Remove(field.Declaration, rewriter);
                    //rewriter.Remove(target.Declaration);
                    continue;
                }

                rewriter.Rename(field.Declaration, field.NewFieldName);
                rewriter.SetVariableVisiblity(field.Declaration, Accessibility.Private.TokenString());
                rewriter.MakeImplicitDeclarationTypeExplicit(field.Declaration);
            }
        }

        private void InsertNewContent(EncapsulateFieldModel model, IExecutableRewriteSession rewriteSession, bool postPendPreviewMessage = false)
        {
            var rewriter = EncapsulateFieldRewriter.CheckoutModuleRewriter(rewriteSession, _targetQMN);

            LoadNewDeclarationBlocks(model);

            LoadNewPropertyBlocks(model);

            if (postPendPreviewMessage)
            {
                _newContent[NewContentTypes.PostContentMessage].Add("'<===== All Changes above this line =====>");
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
                var udt = new UDTDeclarationGenerator(StateUDTField.AsTypeName);

                udt.AddMembers(model.SelectedFieldCandidates);

                AddCodeBlock(NewContentTypes.TypeDeclarationBlock, udt.TypeDeclarationBlock(_indenter));
                AddCodeBlock(NewContentTypes.DeclarationBlock, udt.FieldDeclarationBlock(StateUDTField.NewFieldName));
                return;
            }

            //New field declarations created here were removed from their list within ModifyFields(...)
            var fieldsRequiringNewDeclaration = model.SelectedFieldCandidates
                .Where(field => field.Declaration.IsDeclaredInList()
                                    && field.Declaration.Accessibility != Accessibility.Private);

            foreach (var field in fieldsRequiringNewDeclaration)
            {
                var targetIdentifier = field.Declaration.Context.GetText().Replace(field.IdentifierName, field.NewFieldName);
                var newField = field.Declaration.IsTypeSpecified
                    ? $"{Tokens.Private} {targetIdentifier}"
                    : $"{Tokens.Private} {targetIdentifier} {Tokens.As} {field.Declaration.AsTypeName}";

                AddCodeBlock(NewContentTypes.DeclarationBlock, newField);
            }
        }

        private void LoadNewPropertyBlocks(EncapsulateFieldModel model)
        {
            var propertyGenerationSpecs = model.SelectedFieldCandidates
                                                .SelectMany(f => f.PropertyGenerationSpecs);

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

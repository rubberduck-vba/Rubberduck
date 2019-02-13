using System;
using System.Linq;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoring : IRefactoring
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IParseManager _parseManager;
        private readonly IMessageBox _messageBox;
        private readonly Func<ExtractInterfaceModel, IDisposalActionContainer<IExtractInterfacePresenter>> _presenterFactory;
        private readonly IRewritingManager _rewritingManager;
        private readonly ISelectionService _selectionService;
        private ExtractInterfaceModel _model;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ExtractInterfaceRefactoring(IDeclarationFinderProvider declarationFinderProvider, IParseManager parseManager, IMessageBox messageBox, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _parseManager = parseManager;
            _rewritingManager = rewritingManager;
            _messageBox = messageBox;
            _selectionService = selectionService;
            _presenterFactory = ((model) => DisposalActionContainer.Create(factory.Create<IExtractInterfacePresenter, ExtractInterfaceModel>(model), factory.Release));
        }

        private ExtractInterfaceModel InitializeModel()
        {
            var activeSelection = _selectionService.ActiveSelection();
            if (!activeSelection.HasValue)
            {
                return null;
            }

            return new ExtractInterfaceModel(_declarationFinderProvider, activeSelection.Value);
        }

        public void Refactor()
        {
            _model = InitializeModel();

            if (_model == null)
            {
                return;
            }

            using (var presenterContainer = _presenterFactory(_model))
            {
                var presenter = presenterContainer.Value;
                if (presenter == null)
                {
                    return;
                }

                _model = presenter.Show();
                if (_model == null)
                {
                    return;
                }

                AddInterface();
            }
        }

        public void Refactor(QualifiedSelection target)
        {
            if (!_selectionService.TrySetActiveSelection(target))
            {
                return;
            }

            Refactor();
        }

        public void Refactor(Declaration target)
        {
            if (target == null)
            {
                return;
            }

            Refactor(target.QualifiedSelection);
        }

        private void AddInterface()
        {
            //We need to suspend here since adding the interface and rewriting will both trigger a reparse.
            var suspendResult = _parseManager.OnSuspendParser(this, new[] {ParserState.Ready}, AddInterfaceInternal);
            if (suspendResult != SuspensionResult.Completed)
            {
                _logger.Warn("Extract interface failed.");
            }
        }

        private void AddInterfaceInternal()
        {
            var targetProject = _model.TargetDeclaration.Project;
            if (targetProject == null)
            {
                return; //The target project is not available.
            }

            AddInterfaceClass(_model.TargetDeclaration, _model.InterfaceName, GetInterfaceModuleBody());

            var rewriteSession = _rewritingManager.CheckOutCodePaneSession();
            var rewriter = rewriteSession.CheckOutModuleRewriter(_model.TargetDeclaration.QualifiedModuleName);

            var firstNonFieldMember = _declarationFinderProvider.DeclarationFinder.Members(_model.TargetDeclaration)
                                            .OrderBy(o => o.Selection)
                                            .First(m => ExtractInterfaceModel.MemberTypes.Contains(m.DeclarationType));
            rewriter.InsertBefore(firstNonFieldMember.Context.Start.TokenIndex, $"Implements {_model.InterfaceName}{Environment.NewLine}{Environment.NewLine}");

            AddInterfaceMembersToClass(rewriter);

            rewriteSession.TryRewrite();
        }

        private void AddInterfaceClass(Declaration implementingClass, string interfaceName, string interfaceBody)
        {
            var targetProject = implementingClass.Project;
            using (var components = targetProject.VBComponents)
            {
                using (var interfaceComponent = components.Add(ComponentType.ClassModule))
                {
                    using (var interfaceModule = interfaceComponent.CodeModule)
                    {
                        interfaceComponent.Name = interfaceName;

                        var optionPresent = interfaceModule.CountOfLines > 1;
                        if (!optionPresent)
                        {
                            interfaceModule.InsertLines(1, $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}");
                        }
                        interfaceModule.InsertLines(3, interfaceBody);
                    }
                }
            }
        }

        private void AddInterfaceMembersToClass(IModuleRewriter rewriter)
        {
            var implementInterfaceRefactoring = new ImplementInterfaceRefactoring(_declarationFinderProvider, _messageBox, _rewritingManager, _selectionService);
            implementInterfaceRefactoring.Refactor(_model.SelectedMembers.Select(m => m.Member).ToList(), rewriter, _model.InterfaceName);
        }

        private string GetInterfaceModuleBody()
        {
            return string.Join(Environment.NewLine, _model.SelectedMembers.Select(m => m.Body));
        }
    }
}

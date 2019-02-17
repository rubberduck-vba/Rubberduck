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
    public class ExtractInterfaceRefactoring : InteractiveRefactoringBase<IExtractInterfacePresenter, ExtractInterfaceModel>
    {
        private readonly IDeclarationFinderProvider _declarationFinderProvider;
        private readonly IParseManager _parseManager;
        private readonly IMessageBox _messageBox;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ExtractInterfaceRefactoring(IDeclarationFinderProvider declarationFinderProvider, IParseManager parseManager, IMessageBox messageBox, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService, factory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _parseManager = parseManager;
            _messageBox = messageBox;
        }

        public override void Refactor(QualifiedSelection target)
        {
            Refactor(InitializeModel(target));
        }

        private ExtractInterfaceModel InitializeModel(QualifiedSelection targetSelection)
        {
            return new ExtractInterfaceModel(_declarationFinderProvider, targetSelection);
        }

        protected override void RefactorImpl(IExtractInterfacePresenter presenter)
        {
            AddInterface();
        }

        public override void Refactor(Declaration target)
        {
            Refactor(InitializeModel(target));
        }

        private ExtractInterfaceModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                return null;
            }

            return InitializeModel(target.QualifiedSelection);
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
            var targetProject = Model.TargetDeclaration.Project;
            if (targetProject == null)
            {
                return; //The target project is not available.
            }

            AddInterfaceClass(Model.TargetDeclaration, Model.InterfaceName, GetInterfaceModuleBody());

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            var rewriter = rewriteSession.CheckOutModuleRewriter(Model.TargetDeclaration.QualifiedModuleName);

            var firstNonFieldMember = _declarationFinderProvider.DeclarationFinder.Members(Model.TargetDeclaration)
                                            .OrderBy(o => o.Selection)
                                            .First(m => ExtractInterfaceModel.MemberTypes.Contains(m.DeclarationType));
            rewriter.InsertBefore(firstNonFieldMember.Context.Start.TokenIndex, $"Implements {Model.InterfaceName}{Environment.NewLine}{Environment.NewLine}");

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
            var implementInterfaceRefactoring = new ImplementInterfaceRefactoring(_declarationFinderProvider, _messageBox, RewritingManager, SelectionService);
            implementInterfaceRefactoring.Refactor(Model.SelectedMembers.Select(m => m.Member).ToList(), rewriter, Model.InterfaceName);
        }

        private string GetInterfaceModuleBody()
        {
            return string.Join(Environment.NewLine, Model.SelectedMembers.Select(m => m.Body));
        }
    }
}

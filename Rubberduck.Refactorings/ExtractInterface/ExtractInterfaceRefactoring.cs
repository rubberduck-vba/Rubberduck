using System;
using System.Linq;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
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

        private readonly ImplementInterfaceRefactoring _implementInterfaceRefactoring;

        private readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public ExtractInterfaceRefactoring(IDeclarationFinderProvider declarationFinderProvider, IParseManager parseManager, IRefactoringPresenterFactory factory, IRewritingManager rewritingManager, ISelectionService selectionService)
        :base(rewritingManager, selectionService, factory)
        {
            _declarationFinderProvider = declarationFinderProvider;
            _parseManager = parseManager;

            _implementInterfaceRefactoring = new ImplementInterfaceRefactoring(_declarationFinderProvider, RewritingManager, SelectionService);
        }

        protected override Declaration FindTargetDeclaration(QualifiedSelection targetSelection)
        {
            var candidates = _declarationFinderProvider.DeclarationFinder
                .Members(targetSelection.QualifiedName)
                .Where(item => ModuleTypes.Contains(item.DeclarationType));

            return candidates.SingleOrDefault(item =>
                item.QualifiedSelection.QualifiedName.Equals(targetSelection.QualifiedName));
        }

        protected override ExtractInterfaceModel InitializeModel(Declaration target)
        {
            if (target == null)
            {
                throw new TargetDeclarationIsNullException();
            }

            if (!ModuleTypes.Contains(target.DeclarationType))
            {
                throw new InvalidDeclarationTypeException(target);
            }

            return new ExtractInterfaceModel(_declarationFinderProvider, target);
        }

        protected override void RefactorImpl(ExtractInterfaceModel model)
        {
            AddInterface(model);
        }

        private void AddInterface(ExtractInterfaceModel model)
        {
            //We need to suspend here since adding the interface and rewriting will both trigger a reparse.
            var suspendResult = _parseManager.OnSuspendParser(this, new[] {ParserState.Ready}, () => AddInterfaceInternal(model));
            if (suspendResult != SuspensionResult.Completed)
            {
                _logger.Warn("Extract interface failed.");
            }
        }

        private void AddInterfaceInternal(ExtractInterfaceModel model)
        {
            var targetProject = model.TargetDeclaration.Project;
            if (targetProject == null)
            {
                return; //The target project is not available.
            }

            AddInterfaceClass(model.TargetDeclaration, model.InterfaceName, GetInterfaceModuleBody(model));

            var rewriteSession = RewritingManager.CheckOutCodePaneSession();
            var rewriter = rewriteSession.CheckOutModuleRewriter(model.TargetDeclaration.QualifiedModuleName);

            var firstNonFieldMember = _declarationFinderProvider.DeclarationFinder.Members(model.TargetDeclaration)
                                            .OrderBy(o => o.Selection)
                                            .First(m => ExtractInterfaceModel.MemberTypes.Contains(m.DeclarationType));
            rewriter.InsertBefore(firstNonFieldMember.Context.Start.TokenIndex, $"Implements {model.InterfaceName}{Environment.NewLine}{Environment.NewLine}");

            AddInterfaceMembersToClass(model, rewriter);

            if (!rewriteSession.TryRewrite())
            {
                throw new RewriteFailedException(rewriteSession);
            }
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

        private void AddInterfaceMembersToClass(ExtractInterfaceModel model, IModuleRewriter rewriter)
        {
            _implementInterfaceRefactoring.Refactor(model.SelectedMembers.Select(m => m.Member).ToList(), rewriter, model.InterfaceName);
        }

        private string GetInterfaceModuleBody(ExtractInterfaceModel model)
        {
            return string.Join(Environment.NewLine, model.SelectedMembers.Select(m => m.Body));
        }

        private static readonly DeclarationType[] ModuleTypes =
        {
            DeclarationType.ClassModule,
            DeclarationType.Document,
            DeclarationType.UserForm
        };
    }
}

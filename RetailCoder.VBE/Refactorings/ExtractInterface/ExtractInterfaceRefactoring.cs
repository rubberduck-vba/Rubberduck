using System;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.PostProcessing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory<IExtractInterfacePresenter> _factory;
        private ExtractInterfaceModel _model;

        public ExtractInterfaceRefactoring(IVBE vbe, IMessageBox messageBox, IRefactoringPresenterFactory<IExtractInterfacePresenter> factory)
        {
            _vbe = vbe;
            _messageBox = messageBox;
            _factory = factory;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            if (presenter == null)
            {
                return;
            }

            _model = presenter.Show();
            if (_model == null)
            {
                return;
            }

            var pane = _vbe.ActiveCodePane;
            QualifiedSelection? oldSelection;
            if (!pane.IsWrappingNullReference)
            {
                var module = pane.CodeModule;
                {
                    oldSelection = module.GetQualifiedSelection();
                }
            }
            else
            {
                return;
            }

            AddInterface();

            if (oldSelection.HasValue)
            {
                pane.Selection = oldSelection.Value.Selection;
            }

            _model.State.OnParseRequested(this);
        }

        public void Refactor(QualifiedSelection target)
        {
            var pane = _vbe.ActiveCodePane;
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }
                pane.Selection = target.Selection;
            }
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            var pane = _vbe.ActiveCodePane;
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }
                pane.Selection = target.QualifiedSelection.Selection;
            }
            Refactor();
        }

        private void AddInterface()
        {
            var rewriter = _model.State.GetRewriter(_model.TargetDeclaration);

            var firstNonFieldMember = _model.State.DeclarationFinder.Members(_model.TargetDeclaration)
                                            .OrderBy(o => o.Selection)
                                            .First(m => ExtractInterfaceModel.MemberTypes.Contains(m.DeclarationType));
            rewriter.InsertBefore(firstNonFieldMember.Context.Start.TokenIndex, $"Implements {_model.InterfaceName}{Environment.NewLine}{Environment.NewLine}");

            AddInterfaceMembersToClass(rewriter);

            var components = _model.TargetDeclaration.Project.VBComponents;
            var interfaceComponent = components.Add(ComponentType.ClassModule);
            var interfaceModule = interfaceComponent.CodeModule;
            interfaceComponent.Name = _model.InterfaceName;

            var optionPresent = interfaceModule.CountOfLines > 1;
            if (!optionPresent)
            {
                interfaceModule.InsertLines(1, Tokens.Option + ' ' + Tokens.Explicit + Environment.NewLine);
            }
            interfaceModule.InsertLines(3, GetInterfaceModuleBody());
        }

        private void AddInterfaceMembersToClass(IModuleRewriter rewriter)
        {
            var implementInterfaceRefactoring = new ImplementInterfaceRefactoring(_vbe, _model.State, _messageBox);
            implementInterfaceRefactoring.Refactor(_model.Members.Where(m => m.IsSelected).Select(m => m.Member).ToList(), rewriter, _model.InterfaceName);
        }

        private string GetInterfaceModuleBody()
        {
            return string.Join(Environment.NewLine, _model.Members.Where(m => m.IsSelected).Select(m => m.Body));
        }
    }
}

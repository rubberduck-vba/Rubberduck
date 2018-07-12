using System;
using System.Linq;
using Rubberduck.Interaction;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.ImplementInterface;
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

            using (var pane = _vbe.ActiveCodePane)
            {
                if (pane.IsWrappingNullReference)
                {
                    return;
                }

                var oldSelection = pane.GetQualifiedSelection();

                AddInterface();

                if (oldSelection.HasValue)
                {
                    pane.Selection = oldSelection.Value.Selection;
                }
            }

            _model.State.OnParseRequested(this);
        }

        public void Refactor(QualifiedSelection target)
        {
            using (var pane = _vbe.ActiveCodePane)
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
            using (var pane = _vbe.ActiveCodePane)
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
            var targetProject = _model.TargetDeclaration.Project;
            if (targetProject == null)
            {
                return; //The target project is not available.
            }

            var rewriter = _model.State.GetRewriter(_model.TargetDeclaration);

            var firstNonFieldMember = _model.State.DeclarationFinder.Members(_model.TargetDeclaration)
                                            .OrderBy(o => o.Selection)
                                            .First(m => ExtractInterfaceModel.MemberTypes.Contains(m.DeclarationType));
            rewriter.InsertBefore(firstNonFieldMember.Context.Start.TokenIndex, $"Implements {_model.InterfaceName}{Environment.NewLine}{Environment.NewLine}");

            AddInterfaceMembersToClass(rewriter);

            using (var components = targetProject.VBComponents)
            {
                using (var interfaceComponent = components.Add(ComponentType.ClassModule))
                {
                    using (var interfaceModule = interfaceComponent.CodeModule)
                    {
                        interfaceComponent.Name = _model.InterfaceName;

                        var optionPresent = interfaceModule.CountOfLines > 1;
                        if (!optionPresent)
                        {
                            interfaceModule.InsertLines(1, $"{Tokens.Option} {Tokens.Explicit}{Environment.NewLine}");
                        }
                        interfaceModule.InsertLines(3, GetInterfaceModuleBody());
                    }
                }
            }
        }

        private void AddInterfaceMembersToClass(IModuleRewriter rewriter)
        {
            var implementInterfaceRefactoring = new ImplementInterfaceRefactoring(_vbe, _model.State, _messageBox);
            implementInterfaceRefactoring.Refactor(_model.Members.Select(m => m.Member).ToList(), rewriter, _model.InterfaceName);
        }

        private string GetInterfaceModuleBody()
        {
            return string.Join(Environment.NewLine, _model.Members.Select(m => m.Body));
        }
    }
}

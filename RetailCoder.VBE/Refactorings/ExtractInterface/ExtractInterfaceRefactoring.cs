using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.UI;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoring : IRefactoring
    {
        private readonly RubberduckParserState _state;
        private readonly IRefactoringPresenterFactory<ExtractInterfacePresenter> _factory;
        private readonly IActiveCodePaneEditor _editor;
        private ExtractInterfaceModel _model;

        public ExtractInterfaceRefactoring(RubberduckParserState state, IRefactoringPresenterFactory<ExtractInterfacePresenter> factory,
            IActiveCodePaneEditor editor)
        {
            _state = state;
            _factory = factory;
            _editor = editor;
        }

        public void Refactor()
        {
            var presenter = _factory.Create();
            if (presenter == null)
            {
                return;
            }

            _model = presenter.Show();

            if (_model == null) { return; }

            AddInterface();
        }

        public void Refactor(QualifiedSelection target)
        {
            _editor.SetSelection(target);
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            _editor.SetSelection(target.QualifiedSelection);
            Refactor();
        }

        private void AddInterface()
        {
            var interfaceComponent = _model.TargetDeclaration.Project.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
            interfaceComponent.Name = _model.InterfaceName;

            _editor.InsertLines(1, GetInterface());

            var module = _model.TargetDeclaration.QualifiedSelection.QualifiedName.Component.CodeModule;

            var implementsLine = module.CountOfDeclarationLines + 1;
            module.InsertLines(implementsLine, "Implements " + _model.InterfaceName);

            _state.RequestParse(ParserState.Ready);
            var qualifiedSelection = new QualifiedSelection(_model.TargetDeclaration.QualifiedSelection.QualifiedName,
                new Selection(implementsLine, 1, implementsLine, 1));

            var implementInterfaceRefactoring = new ImplementInterfaceRefactoring(_state, _editor, new MessageBox());
            implementInterfaceRefactoring.Refactor(qualifiedSelection);
        }

        private string GetInterface()
        {
            return "Option Explicit" + Environment.NewLine + string.Join(Environment.NewLine, _model.Members.Where(m => m.IsSelected));
        }
    }
}
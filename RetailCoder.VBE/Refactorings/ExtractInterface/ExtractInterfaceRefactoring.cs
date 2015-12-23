using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoring : IRefactoring
    {
        private readonly IRefactoringPresenterFactory<ExtractInterfacePresenter> _factory;
        private readonly IActiveCodePaneEditor _editor;
        private ExtractInterfaceModel _model;

        public ExtractInterfaceRefactoring(IRefactoringPresenterFactory<ExtractInterfacePresenter> factory,
            IActiveCodePaneEditor editor)
        {
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
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            Refactor();
        }

        private void AddInterface()
        {
            var interfaceComponent = _model.TargetDeclaration.Project.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
            interfaceComponent.Name = _model.InterfaceName;

            _editor.InsertLines(1, GetInterface());
        }

        private string GetInterface()
        {
            return string.Join(Environment.NewLine, _model.Members.Where(m => m.IsSelected));
        }
    }
}
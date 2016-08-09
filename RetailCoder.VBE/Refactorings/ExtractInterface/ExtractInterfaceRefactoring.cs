using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using NLog;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoring : IRefactoring
    {
        private readonly VBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory<IExtractInterfacePresenter> _factory;
        private ExtractInterfaceModel _model;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public ExtractInterfaceRefactoring(VBE vbe, RubberduckParserState state, IMessageBox messageBox, IRefactoringPresenterFactory<IExtractInterfacePresenter> factory)
        {
            _vbe = vbe;
            _state = state;
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

            QualifiedSelection? oldSelection = null;
            if (_vbe.ActiveCodePane != null)
            {
                oldSelection = _vbe.ActiveCodePane.CodeModule.GetSelection();
            }

            AddInterface();

            if (oldSelection.HasValue)
            {
                oldSelection.Value.QualifiedName.Component.CodeModule.SetSelection(oldSelection.Value.Selection);
                oldSelection.Value.QualifiedName.Component.CodeModule.CodePane.ForceFocus();
            }

            _state.OnParseRequested(this);
        }

        public void Refactor(QualifiedSelection target)
        {
            _vbe.ActiveCodePane.CodeModule.SetSelection(target);
            Refactor();
        }

        public void Refactor(Declaration target)
        {
            _vbe.ActiveCodePane.CodeModule.SetSelection(target.QualifiedSelection);
            Refactor();
        }

        private void AddInterface()
        {
            var interfaceComponent = _model.TargetDeclaration.Project.VBComponents.Add(vbext_ComponentType.vbext_ct_ClassModule);
            interfaceComponent.Name = _model.InterfaceName;

            interfaceComponent.CodeModule.InsertLines(1, Tokens.Option + ' ' + Tokens.Explicit + Environment.NewLine);
            interfaceComponent.CodeModule.InsertLines(3, GetInterfaceModuleBody());

            var module = _model.TargetDeclaration.QualifiedSelection.QualifiedName.Component.CodeModule;

            _insertionLine = module.CountOfDeclarationLines + 1;
            module.InsertLines(_insertionLine, Tokens.Implements + ' ' + _model.InterfaceName + Environment.NewLine);

            _state.StateChanged += _state_StateChanged;
            _state.OnParseRequested(this);
        }

        private int _insertionLine;
        private void _state_StateChanged(object sender, EventArgs e)
        {
            if (_state.Status != ParserState.Ready)
            {
                return;
            }

            _state.StateChanged -= _state_StateChanged;
            var qualifiedSelection = new QualifiedSelection(_model.TargetDeclaration.QualifiedSelection.QualifiedName, new Selection(_insertionLine, 1, _insertionLine, 1));
            _vbe.ActiveCodePane.CodeModule.SetSelection(qualifiedSelection);

            var implementInterfaceRefactoring = new ImplementInterfaceRefactoring(_vbe, _state, _messageBox);
            implementInterfaceRefactoring.Refactor(qualifiedSelection);
        }

        private string GetInterfaceModuleBody()
        {
            return string.Join(Environment.NewLine, _model.Members.Where(m => m.IsSelected).Select(m => m.Body));
        }
    }
}

using System;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ImplementInterface;
using Rubberduck.UI;
using Rubberduck.VBEditor;
using NLog;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class ExtractInterfaceRefactoring : IRefactoring
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory<IExtractInterfacePresenter> _factory;
        private ExtractInterfaceModel _model;
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        public ExtractInterfaceRefactoring(IVBE vbe, RubberduckParserState state, IMessageBox messageBox, IRefactoringPresenterFactory<IExtractInterfacePresenter> factory)
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

            var pane = _vbe.ActiveCodePane;
            {
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
            }
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
            var components = _model.TargetDeclaration.Project.VBComponents;
            var interfaceComponent = components.Add(ComponentType.ClassModule);
            var interfaceModule = interfaceComponent.CodeModule;
            {
                interfaceComponent.Name = _model.InterfaceName;

                var optionPresent = interfaceModule.CountOfLines > 1;
                if (!optionPresent)
                {
                    interfaceModule.InsertLines(1, Tokens.Option + ' ' + Tokens.Explicit + Environment.NewLine);
                }
                interfaceModule.InsertLines(3, GetInterfaceModuleBody());

                var module = _model.TargetDeclaration.QualifiedSelection.QualifiedName.Component.CodeModule;
                {
                    _insertionLine = module.CountOfDeclarationLines + 1;
                    module.InsertLines(_insertionLine, Tokens.Implements + ' ' + _model.InterfaceName + Environment.NewLine);

                    _state.StateChanged += _state_StateChanged;
                    _state.OnParseRequested(this);
                }
            }
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
            var pane = _vbe.ActiveCodePane;
            {
                pane.Selection = qualifiedSelection.Selection;
            }

            var implementInterfaceRefactoring = new ImplementInterfaceRefactoring(_vbe, _state, _messageBox);
            implementInterfaceRefactoring.Refactor(qualifiedSelection);
        }

        private string GetInterfaceModuleBody()
        {
            return string.Join(Environment.NewLine, _model.Members.Where(m => m.IsSelected).Select(m => m.Body));
        }
    }
}

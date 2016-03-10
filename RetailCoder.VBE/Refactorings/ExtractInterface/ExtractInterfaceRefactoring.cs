﻿using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Grammar;
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
        private readonly IMessageBox _messageBox;
        private readonly IRefactoringPresenterFactory<ExtractInterfacePresenter> _factory;
        private readonly IActiveCodePaneEditor _editor;
        private ExtractInterfaceModel _model;

        public ExtractInterfaceRefactoring(RubberduckParserState state, IMessageBox messageBox, IRefactoringPresenterFactory<ExtractInterfacePresenter> factory,
            IActiveCodePaneEditor editor)
        {
            _state = state;
            _messageBox = messageBox;
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

            if (_model == null)
            {
                return;
            }

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

            _editor.InsertLines(1, Tokens.Option + ' ' + Tokens.Explicit + Environment.NewLine);
            _editor.InsertLines(3, GetInterfaceModuleBody());

            var module = _model.TargetDeclaration.QualifiedSelection.QualifiedName.Component.CodeModule;

            _insertionLine = module.CountOfDeclarationLines + 1;
            module.InsertLines(_insertionLine, Tokens.Implements + ' ' + _model.InterfaceName + Environment.NewLine);

            _state.StateChanged += _state_StateChanged;
            _state.OnParseRequested(this);
        }

        private int _insertionLine;
        private void _state_StateChanged(object sender, ParserStateEventArgs e)
        {
            if (e.State != ParserState.Ready)
            {
                return;
            }

            var qualifiedSelection = new QualifiedSelection(_model.TargetDeclaration.QualifiedSelection.QualifiedName, new Selection(_insertionLine, 1, _insertionLine, 1));
            _editor.SetSelection(qualifiedSelection);

            var implementInterfaceRefactoring = new ImplementInterfaceRefactoring(_state, _editor, _messageBox);
            implementInterfaceRefactoring.Refactor(qualifiedSelection);

            _state.StateChanged -= _state_StateChanged;
        }

        private string GetInterfaceModuleBody()
        {
            return string.Join(Environment.NewLine, _model.Members.Where(m => m.IsSelected).Select(m => m.Body));
        }
    }
}
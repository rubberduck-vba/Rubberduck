using System;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.DisposableWrappers.Office.Core;
using Rubberduck.VBEditor.DisposableWrappers.VBA;
using Rubberduck.VBEditor.Extensions;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RubberduckCommandBar : IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly VBE _vbe;
        private readonly ISinks _sinks;
        private readonly IShowParserErrorsCommand _command;

        private CommandBarButton _refreshButton;
        private CommandBarButton _statusButton;
        private CommandBarButton _selectionButton;
        private CommandBar _commandbar;

        public RubberduckCommandBar(RubberduckParserState state, VBE vbe, ISinks sinks, IShowParserErrorsCommand command)
        {
            _state = state;
            _vbe = vbe;
            _sinks = sinks;
            _command = command;
            _state.StateChanged += State_StateChanged;
            Initialize();

            _sinks.ProjectRemoved += ProjectRemoved;
            _sinks.ComponentActivated += ComponentActivated;
            _sinks.ComponentSelected += ComponentSelected;
        }

        private void ProjectRemoved(object sender, IProjectEventArgs e)
        {
            SetSelectionText();
        }

        private void ComponentActivated(object sender, IComponentEventArgs e)
        {
            SetSelectionText();
        }

        private void ComponentSelected(object sender, IComponentEventArgs e)
        {
            SetSelectionText();
        }

        private void SetSelectionText()
        {
            var selectedDeclaration = _vbe.ActiveCodePane != null
                            ? _state.FindSelectedDeclaration(_vbe.ActiveCodePane)
                            : null;

            SetSelectionText(selectedDeclaration);
        }

        private void _statusButton_Click(object sender, CommandBarButtonClickEventArgs e)
        {
            if (_state.Status == ParserState.Error)
            {
                _command.Execute(null);
            }
            e.Cancel = true;
        }

        public void SetStatusText(string value = null)
        {
            var text = value ?? RubberduckUI.ResourceManager.GetString("ParserState_" + _state.Status, Settings.Settings.Culture);
            UiDispatcher.Invoke(() => _statusButton.Caption = text);
        }

        public void SetSelectionText(Declaration declaration)
        {
            if (declaration == null && _vbe.ActiveCodePane != null)
            {
                var selection = _vbe.ActiveCodePane.GetQualifiedSelection();
                if (selection.HasValue) { SetSelectionText(selection.Value); }
                _selectionButton.TooltipText = _selectionButton.Caption;
            }
            else if (declaration == null && _vbe.ActiveCodePane == null)
            {
                UiDispatcher.Invoke(() => _selectionButton.Caption = string.Empty);
            }
            else if (declaration != null && !declaration.IsBuiltIn && declaration.DeclarationType != DeclarationType.ClassModule && declaration.DeclarationType != DeclarationType.ProceduralModule)
            {
                var typeName = declaration.HasTypeHint
                    ? Declaration.TypeHintToTypeName[declaration.TypeHint]
                    : declaration.AsTypeName;

                _selectionButton.Caption = string.Format("{0}|{1}: {2} ({3}{4})",
                    declaration.QualifiedSelection.Selection,
                    declaration.QualifiedName.QualifiedModuleName,
                    declaration.IdentifierName,
                    RubberduckUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType, Settings.Settings.Culture),
                    string.IsNullOrEmpty(declaration.AsTypeName) ? string.Empty : ": " + typeName);

                _selectionButton.TooltipText = string.IsNullOrEmpty(declaration.DescriptionString)
                    ? _selectionButton.Caption
                    : declaration.DescriptionString;
            }
            else if (declaration != null)
            {
                // todo: confirm this is what we want, and then refator
                var selection = _vbe.ActiveCodePane.GetQualifiedSelection();
                if (selection.HasValue)
                {
                    var typeName = declaration.HasTypeHint
                        ? Declaration.TypeHintToTypeName[declaration.TypeHint]
                        : declaration.AsTypeName;

                    _selectionButton.Caption = string.Format("{0}|{1}: {2} ({3}{4})",
                        selection.Value.Selection,
                        declaration.QualifiedName.QualifiedModuleName,
                        declaration.IdentifierName,
                        RubberduckUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType, Settings.Settings.Culture),
                        string.IsNullOrEmpty(declaration.AsTypeName) ? string.Empty : ": " + typeName);
                }
                _selectionButton.TooltipText = string.IsNullOrEmpty(declaration.DescriptionString)
                    ? _selectionButton.Caption
                    : declaration.DescriptionString;
            }
        }

        private void SetSelectionText(QualifiedSelection selection)
        {
            UiDispatcher.Invoke(() => _selectionButton.Caption = selection.ToString());
        }

        private void State_StateChanged(object sender, EventArgs e)
        {
            if (_state.Status != ParserState.ResolvedDeclarations)
            {
                SetStatusText(RubberduckUI.ResourceManager.GetString("ParserState_" + _state.Status, Settings.Settings.Culture));
            }
        }

        public event EventHandler Refresh;

        private void OnRefresh()
        {
            var handler = Refresh;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }

        private void Initialize()
        {
            _commandbar = _vbe.CommandBars.Add("Rubberduck", CommandBarPosition.Top);

            _refreshButton = CommandBarButton.FromCommandBarControl(_commandbar.Controls.Add(ControlType.Button));
            _refreshButton.Picture = Resources.arrow_circle_double;
            _refreshButton.Mask = Resources.arrow_circle_double_mask;
            _refreshButton.Style = ButtonStyle.Icon;
            _refreshButton.Tag = "Refresh";
            _refreshButton.TooltipText = RubberduckUI.RubberduckCommandbarRefreshButtonTooltip;
            _refreshButton.Click += refreshButton_Click;
            _refreshButton.ApplyIcon();

            _statusButton = CommandBarButton.FromCommandBarControl(_commandbar.Controls.Add(ControlType.Button));
            _statusButton.Style = ButtonStyle.Caption;
            _statusButton.Tag = "Status";
            _statusButton.Click += _statusButton_Click;

            _selectionButton = CommandBarButton.FromCommandBarControl(_commandbar.Controls.Add(ControlType.Button));
            _selectionButton.Style = ButtonStyle.Caption;
            _selectionButton.BeginsGroup = true;
            _selectionButton.IsEnabled = false;

            _commandbar.IsVisible = true;
            _sinks.Start();
        }

        private void refreshButton_Click(object sender, CommandBarButtonClickEventArgs e)
        {
            OnRefresh();
            e.Cancel = true;
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            _state.StateChanged -= State_StateChanged;

            _sinks.ProjectRemoved -= ProjectRemoved;
            _sinks.ComponentActivated -= ComponentActivated;
            _sinks.ComponentSelected -= ComponentSelected;

            _refreshButton.Delete();
            _selectionButton.Delete();
            _statusButton.Delete();
            _commandbar.Delete();
            _commandbar.Release();
            _isDisposed = true;
        }
    }
}

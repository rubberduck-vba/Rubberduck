using System;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Properties;
using Rubberduck.UI.Command.MenuItems.ParentMenus;
using Rubberduck.VBEditor;
using ParserState = Rubberduck.Parsing.VBA.ParserState;
using NLog;

namespace Rubberduck.UI.Command.MenuItems
{
    public class RubberduckCommandBar : IDisposable
    {
        private readonly RubberduckParserState _state;
        private readonly VBE _vbe;
        private readonly IShowParserErrorsCommand _command;

        private CommandBarButton _refreshButton;
        private CommandBarButton _statusButton;
        private CommandBarButton _selectionButton;
        private CommandBar _commandbar;
        private static readonly Logger _logger = LogManager.GetCurrentClassLogger();

        public RubberduckCommandBar(RubberduckParserState state, VBE vbe, IShowParserErrorsCommand command)
        {
            _state = state;
            _vbe = vbe;
            _command = command;
            _state.StateChanged += State_StateChanged;
            Initialize();
        }

        private void _statusButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            if (_state.Status == ParserState.Error)
            {
                _command.Execute(null);
            }
        }

        public void SetStatusText(string value = null)
        {
            var text = value ?? RubberduckUI.ResourceManager.GetString("ParserState_" + _state.Status);
            _logger.Debug("RubberduckCommandBar status text changes to '{0}'.", text);
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
            else if (declaration != null && !declaration.IsBuiltIn && declaration.DeclarationType != DeclarationType.ClassModule && declaration.DeclarationType != DeclarationType.ProceduralModule)
            {
                _selectionButton.Caption = string.Format("{0}|{1}: {2} ({3}{4})",
                    declaration.QualifiedSelection.Selection,
                    declaration.QualifiedName.QualifiedModuleName,
                    declaration.IdentifierName,
                    RubberduckUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType),
                    string.IsNullOrEmpty(declaration.AsTypeName) ? string.Empty : ": " + declaration.AsTypeName);
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
                    _selectionButton.Caption = string.Format("{0}|{1}: {2} ({3}{4})",
                        selection.Value.Selection,
                        declaration.QualifiedName.QualifiedModuleName,
                        declaration.IdentifierName,
                        RubberduckUI.ResourceManager.GetString("DeclarationType_" + declaration.DeclarationType),
                    string.IsNullOrEmpty(declaration.AsTypeName) ? string.Empty : ": " + declaration.AsTypeName);
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
            _logger.Debug("RubberduckCommandBar handles StateChanged...");
            
            if (_state.Status != ParserState.ResolvedDeclarations)
            {
                SetStatusText(RubberduckUI.ResourceManager.GetString("ParserState_" + _state.Status));
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

        public void Initialize()
        {
            _commandbar = _vbe.CommandBars.Add("Rubberduck", MsoBarPosition.msoBarTop, false, true);

            _refreshButton = (CommandBarButton)_commandbar.Controls.Add(MsoControlType.msoControlButton);
            ParentMenuItemBase.SetButtonImage(_refreshButton, Resources.arrow_circle_double, Resources.arrow_circle_double_mask);
            _refreshButton.Style = MsoButtonStyle.msoButtonIcon;
            _refreshButton.Tag = "Refresh";
            _refreshButton.TooltipText =RubberduckUI.RubberduckCommandbarRefreshButtonTooltip;
            _refreshButton.Click += refreshButton_Click;

            _statusButton = (CommandBarButton)_commandbar.Controls.Add(MsoControlType.msoControlButton);
            _statusButton.Style = MsoButtonStyle.msoButtonCaption;
            _statusButton.Tag = "Status";
            _statusButton.Click += _statusButton_Click;

            _selectionButton = (CommandBarButton)_commandbar.Controls.Add(MsoControlType.msoControlButton);
            _selectionButton.Style = MsoButtonStyle.msoButtonCaption;
            _selectionButton.BeginGroup = true;
            _selectionButton.Enabled = false;

            _commandbar.Visible = true;
        }

        private void refreshButton_Click(CommandBarButton Ctrl, ref bool CancelDefault)
        {
            OnRefresh();
        }

        private bool _isDisposed;
        public void Dispose()
        {
            if (_isDisposed)
            {
                return;
            }

            _state.StateChanged -= State_StateChanged;

            _refreshButton.Delete();
            _selectionButton.Delete();
            _statusButton.Delete();
            _commandbar.Delete();

            _isDisposed = true;
        }
    }
}

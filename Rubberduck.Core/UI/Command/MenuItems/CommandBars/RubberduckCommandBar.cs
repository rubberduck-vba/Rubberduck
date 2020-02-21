using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Resources;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class RubberduckCommandBar : AppCommandBarBase, IDisposable
    {
        private readonly IContextFormatter _formatter;
        private readonly RubberduckParserState _state;
        private readonly ISelectionChangeService _selectionService;

        public RubberduckCommandBar(RubberduckParserState state, IEnumerable<ICommandMenuItem> items, IContextFormatter formatter, ISelectionChangeService selectionService, IUiDispatcher uiDispatcher) 
            : base("Rubberduck", CommandBarPosition.Top, items, uiDispatcher)
        {
            _state = state;
            _formatter = formatter;
            _selectionService = selectionService;
           
            _state.StateChangedHighPriority += OnParserStateChanged;
            _state.StatusMessageUpdate += OnParserStatusMessageUpdate;
            _selectionService.SelectionChanged += OnSelectionChange;
        }

        //This class is instantiated early enough that the initial control state isn't ready to be set up.  So... override Initialize
        //and have the state set via the Initialize call.
        public override void Initialize()
        {
            base.Initialize();
            SetStatusLabelCaption(ParserState.Pending);
            EvaluateCanExecute(_state);
        }

        private Declaration _lastDeclaration;
        private ParserState _lastStatus = ParserState.None;
        private void EvaluateCanExecute(RubberduckParserState state, Declaration selected)
        {
            var currentStatus = _state.Status;
            if (_lastStatus == currentStatus && 
                (selected == null || selected.Equals(_lastDeclaration)) &&
                (selected != null || _lastDeclaration == null))
            {
                return;
            }

            _lastStatus = currentStatus;
            _lastDeclaration = selected;
            base.EvaluateCanExecute(state);
        }

        private void OnSelectionChange(object sender, DeclarationChangedEventArgs e)
        {
            var caption = _formatter.Format(e.Declaration, e.MultipleControlsSelected);
            if (string.IsNullOrEmpty(caption))
            {
                //Fallback caption for selections in the Project window.                               
                caption = e.FallbackCaption;
            }

            var refCount = e.Declaration?.References.Count() ?? 0;
            var description = e.Declaration?.DescriptionString ?? string.Empty;
            //& renders the next character as if it was an accelerator.
            SetContextSelectionCaption(caption?.Replace("&", "&&"), refCount, description);
            EvaluateCanExecute(_state, e.Declaration);
        }

        
        private void OnParserStatusMessageUpdate(object sender, RubberduckStatusMessageEventArgs e)
        {
            var message = e.Message;
            if (message == ParserState.LoadingReference.ToString())
            {
                // note: ugly hack to enable Rubberduck.Parsing assembly to do this
                message = RubberduckUI.ParserState_LoadingReference;
            }

            SetStatusLabelCaption(message, _state.ModuleExceptions.Count);            
        }

        private void OnParserStateChanged(object sender, EventArgs e)
        {
            _lastStatus = _state.Status;
            EvaluateCanExecute(_state);    
            SetStatusLabelCaption(_state.Status, _state.ModuleExceptions.Count);                 
        }

        public void SetStatusLabelCaption(ParserState state, int? errorCount = null)
        {
            var caption = RubberduckUI.ResourceManager.GetString("ParserState_" + state, CultureInfo.CurrentUICulture);
            SetStatusLabelCaption(caption, errorCount);
        }

        private void SetStatusLabelCaption(string caption, int? errorCount = null)
        {
            var reparseCommandButton =
                FindChildByTag(typeof(ReparseCommandMenuItem).FullName) as ReparseCommandMenuItem;
            if (reparseCommandButton == null)
            {
                return;
            }

            var showErrorsCommandButton =
                FindChildByTag(typeof(ShowParserErrorsCommandMenuItem).FullName) as ShowParserErrorsCommandMenuItem;
            if (showErrorsCommandButton == null)
            {
                return;
            }

            _uiDispatcher.Invoke(() =>
            {
                try
                {
                    reparseCommandButton.SetCaption(caption);
                    reparseCommandButton.SetToolTip(string.Format(RubberduckUI.ReparseToolTipText, caption));
                    if (errorCount.HasValue && errorCount.Value > 0)
                    {
                        showErrorsCommandButton.SetToolTip(
                            string.Format(RubberduckUI.ParserErrorToolTipText, errorCount.Value));
                    }
                }
                catch (Exception exception)
                {
                    Logger.Error(exception,
                        "Exception thrown trying to set the status label caption on the UI thread.");
                }
            });
            Localize();
        }

        private void SetContextSelectionCaption(string caption, int contextReferenceCount, string description)
        {
            var contextLabel = FindChildByTag(typeof(ContextSelectionLabelMenuItem).FullName) as ContextSelectionLabelMenuItem;
            var contextReferences = FindChildByTag(typeof(ReferenceCounterLabelMenuItem).FullName) as ReferenceCounterLabelMenuItem;
            var contextDescription = FindChildByTag(typeof(ContextDescriptionLabelMenuItem).FullName) as ContextDescriptionLabelMenuItem;

            _uiDispatcher.Invoke(() =>
            {
                try
                {
                    contextLabel?.SetCaption(caption);
                    contextReferences?.SetCaption(contextReferenceCount);
                    contextDescription?.SetCaption(description);
                }
                catch (Exception exception)
                {
                    Logger.Error(exception, "Exception thrown trying to set the context selection caption on the UI thread.");
                }
            });
            Localize();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }

            _selectionService.SelectionChanged -= OnSelectionChange;
            _state.StateChangedHighPriority -= OnParserStateChanged;
            _state.StatusMessageUpdate -= OnParserStatusMessageUpdate;

            RemoveCommandBar();

            _isDisposed = true;
        }
    }

    public enum RubberduckCommandBarItemDisplayOrder
    {
        RequestReparse,
        ShowErrors,
        ContextStatus,
        ContextDescription,
        ContextRefCount,
    }
}
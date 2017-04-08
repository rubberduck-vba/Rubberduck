using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class RubberduckCommandBar : AppCommandBarBase, IDisposable
    {
        private readonly IContextFormatter _formatter;
        private readonly IParseCoordinator _parser;
        private readonly ISelectionChangeService _selectionService;

        public RubberduckCommandBar(IParseCoordinator parser, IEnumerable<ICommandMenuItem> items, IContextFormatter formatter, ISelectionChangeService selectionService) 
            : base("Rubberduck", CommandBarPosition.Top, items)
        {
            _parser = parser;
            _formatter = formatter;
            _selectionService = selectionService;
           
            _parser.State.StateChanged += OnParserStateChanged;
            _parser.State.StatusMessageUpdate += OnParserStatusMessageUpdate;
            _selectionService.SelectionChanged += OnSelectionChange;
        }

        //This class is instantiated early enough that the initial control state isn't ready to be set up.  So... override Initialize
        //and have the state set via the Initialize call.
        public override void Initialize()
        {
            base.Initialize();
            SetStatusLabelCaption(ParserState.Pending);
            EvaluateCanExecute(_parser.State);
        }

        private Declaration _lastDeclaration;
        private ParserState _lastStatus = ParserState.None;
        private void EvaluateCanExecute(RubberduckParserState state, Declaration selected)
        {
            var currentStatus = _parser.State.Status;
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
            var caption = e.ActivePane != null
                ? _formatter.Format(e.ActivePane, e.Declaration)
                : _formatter.Format(e.Declaration, e.MultipleControlsSelected);
           
            if (string.IsNullOrEmpty(caption) && e.VBComponent != null)
            {
                //Fallback caption for selections in the Project window.
                caption = $"{e.VBComponent.ParentProject.Name}.{e.VBComponent.Name} ({e.VBComponent.Type})";
            }

            var refCount = e.Declaration?.References.Count() ?? 0;
            var description = e.Declaration?.DescriptionString ?? string.Empty;
            SetContextSelectionCaption(caption, refCount, description);
            EvaluateCanExecute(_parser.State, e.Declaration);
        }

        
        private void OnParserStatusMessageUpdate(object sender, RubberduckStatusMessageEventArgs e)
        {
            var message = e.Message;
            if (message == ParserState.LoadingReference.ToString())
            {
                // note: ugly hack to enable Rubberduck.Parsing assembly to do this
                message = RubberduckUI.ParserState_LoadingReference;
            }

            SetStatusLabelCaption(message, _parser.State.ModuleExceptions.Count);            
        }

        private void OnParserStateChanged(object sender, EventArgs e)
        {
            _lastStatus = _parser.State.Status;
            EvaluateCanExecute(_parser.State);    
            SetStatusLabelCaption(_parser.State.Status, _parser.State.ModuleExceptions.Count);                 
        }

        public void SetStatusLabelCaption(ParserState state, int? errorCount = null)
        {
            var caption = RubberduckUI.ResourceManager.GetString("ParserState_" + state, CultureInfo.CurrentUICulture);
            SetStatusLabelCaption(caption, errorCount);
        }

        private void SetStatusLabelCaption(string caption, int? errorCount = null)
        {
            var reparseCommandButton = FindChildByTag(typeof(ReparseCommandMenuItem).FullName) as ReparseCommandMenuItem;
            if (reparseCommandButton == null) { return; }

            var showErrorsCommandButton = FindChildByTag(typeof(ShowParserErrorsCommandMenuItem).FullName) as ShowParserErrorsCommandMenuItem;
            if (showErrorsCommandButton == null) { return; }

            UiDispatcher.Invoke(() =>
            {
                reparseCommandButton.SetCaption(caption);
                reparseCommandButton.SetToolTip(string.Format(RubberduckUI.ReparseToolTipText, caption));
                if (errorCount.HasValue && errorCount.Value > 0)
                {
                    showErrorsCommandButton.SetToolTip(string.Format(RubberduckUI.ParserErrorToolTipText, errorCount.Value));
                }
            });
            Localize();
        }

        private void SetContextSelectionCaption(string caption, int contextReferenceCount, string description)
        {
            var contextLabel = FindChildByTag(typeof(ContextSelectionLabelMenuItem).FullName) as ContextSelectionLabelMenuItem;
            var contextReferences = FindChildByTag(typeof(ReferenceCounterLabelMenuItem).FullName) as ReferenceCounterLabelMenuItem;
            var contextDescription = FindChildByTag(typeof(ContextDescriptionLabelMenuItem).FullName) as ContextDescriptionLabelMenuItem;

            UiDispatcher.Invoke(() =>
            {
                contextLabel?.SetCaption(caption);
                contextReferences?.SetCaption(contextReferenceCount);
                contextDescription?.SetCaption(description);
            });
            Localize();
        }

        public void Dispose()
        {
            _selectionService.SelectionChanged -= OnSelectionChange;
            _parser.State.StateChanged -= OnParserStateChanged;
            _parser.State.StatusMessageUpdate -= OnParserStatusMessageUpdate;

            //note: doing this wrecks the teardown process. counter-intuitive? sure. but hey it works.
            //RemoveChildren();
            //Item.Delete();
            //Item.Release(true);
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
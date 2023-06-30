using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
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
            EvaluateCanExecuteAsync(_state, CancellationToken.None);
        }

        private Declaration _lastDeclaration;
        private ParserState _lastStatus = ParserState.None;
        private async Task EvaluateCanExecuteAsync(RubberduckParserState state, Declaration selected, CancellationToken token)
        {
            token.ThrowIfCancellationRequested();
            var currentStatus = _state.Status;
            if (_lastStatus == currentStatus && 
                (selected == null || selected.Equals(_lastDeclaration)) &&
                (selected != null || _lastDeclaration == null))
            {
                return;
            }

            _lastStatus = currentStatus;
            _lastDeclaration = selected;
            await base.EvaluateCanExecuteAsync(state, token);
        }

        private readonly ConcurrentDictionary<string, CancellationTokenSource> _tokenSources = new ConcurrentDictionary<string, CancellationTokenSource>();

        private async void OnSelectionChange(object sender, DeclarationChangedEventArgs e)
        {
            try
            {
                try
                {
                    if (_tokenSources.TryRemove(nameof(OnSelectionChange), out var existing))
                    {
                        existing.Cancel();
                    }
                }
                catch (ObjectDisposedException)
                {
                    Logger.Trace($"CancellationTokenSource was already disposed for {nameof(OnSelectionChange)}.");
                }

                var source = _tokenSources.GetOrAdd(nameof(OnSelectionChange), k => new CancellationTokenSource());
                var token = source.Token;

                Task.Run(async () =>
                    {
                        var caption = await _formatter.FormatAsync(e.Declaration, e.MultipleControlsSelected, token);
                        token.ThrowIfCancellationRequested();

                        var argRefCount = e.Declaration is ParameterDeclaration parameter ? parameter.ArgumentReferences.Count() : 0;
                        var refCount = (e.Declaration?.References.Count() ?? 0) + argRefCount;
                        var description = e.Declaration?.DescriptionString.Trim() ?? string.Empty;
                        token.ThrowIfCancellationRequested();

                        //& renders the next character as if it was an accelerator.
                        SetContextSelectionCaption(caption?.Replace("&", "&&"), refCount, description);
                        token.ThrowIfCancellationRequested();

                        await EvaluateCanExecuteAsync(_state, e.Declaration, token);

                    }, token)
                    .ContinueWith(t =>
                    {
                        try
                        {
                            if (!t.IsCanceled)
                            {
                                source.Dispose();
                            }
                        }
                        catch (Exception exception)
                        {
                            Logger.Trace($"CancellationTokenSource.Dispose() threw an exception for {nameof(OnSelectionChange)}: {exception}");
                        }
                    }, token);
            }
            catch(ObjectDisposedException)
            {
                Logger.Trace($"CancellationTokenSource was already disposed for {nameof(OnSelectionChange)}.");
            }
            catch (OperationCanceledException exception)
            {
                Logger.Info(exception);
            }
            catch (Exception exception)
            {
                Logger.Error(exception);
            }
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

        private async void OnParserStateChanged(object sender, EventArgs e)
        {
            try
            {
                _lastStatus = _state.Status;
                try
                {
                    if (_tokenSources.TryRemove(nameof(OnParserStateChanged), out var existing))
                    {
                        existing.Cancel();
                    }
                }
                catch (ObjectDisposedException)
                {
                    Logger.Trace($"CancellationTokenSource was already disposed for {nameof(OnParserStateChanged)}.");
                }

                var source = _tokenSources.GetOrAdd(nameof(OnParserStateChanged), k => new CancellationTokenSource());
                var token = source.Token;

                await EvaluateCanExecuteAsync(_state, token)
                    .ContinueWith(t =>
                    {
                        try
                        {
                            if (!t.IsCanceled)
                            {
                                source.Dispose();
                            }
                        }
                        catch (Exception exception)
                        {
                            Logger.Trace($"CancellationTokenSource.Dispose() threw an exception for {nameof(OnParserStateChanged)}: {exception}");
                        }
                    }, token);

                SetStatusLabelCaption(_state.Status, _state.ModuleExceptions.Count);
            }
            catch (ObjectDisposedException)
            {
                Logger.Trace($"CancellationTokenSource was already disposed for {nameof(OnParserStateChanged)}.");
            }
            catch (OperationCanceledException exception)
            {
                Logger.Info(exception);
            }
            catch (Exception exception)
            {
                Logger.Error(exception);
            }
        }

        public void SetStatusLabelCaption(ParserState state, int? errorCount = null)
        {
            var caption = RubberduckUI.ResourceManager.GetString("ParserState_" + state, CultureInfo.CurrentUICulture);
            SetStatusLabelCaption(caption, errorCount);
        }

        private void SetStatusLabelCaption(string caption, int? errorCount = null)
        {
            if (!(FindChildByTag(typeof(ReparseCommandMenuItem).FullName) is ReparseCommandMenuItem reparseCommandButton))
            {
                return;
            }

            if (!(FindChildByTag(typeof(ShowParserErrorsCommandMenuItem).FullName) is ShowParserErrorsCommandMenuItem showErrorsCommandButton))
            {
                return;
            }

            _uiDispatcher.Invoke(() =>
            {
                try
                {
                    reparseCommandButton.SetCaption(caption);
                    reparseCommandButton.SetToolTip(string.Format(RubberduckUI.ReparseToolTipText, caption));
                    if (errorCount > 0)
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

            foreach (var source in _tokenSources)
            {
                try
                {
                    source.Value.Dispose();
                }
                catch (Exception exception)
                {
                    Logger.Trace($"Disposing CancellationTokenSource for {nameof(OnParserStateChanged)} threw an exception: {exception}");
                }
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
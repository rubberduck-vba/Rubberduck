using System;
using System.Collections.Concurrent;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace Rubberduck.Parsing.VBA
{
    /// <summary>
    /// Orchestrates parsing tasks.
    /// </summary>
    public class ParseCoordinator
    {
        private readonly ConcurrentDictionary<VBComponent, Task<Action>> _parseTasks =
            new ConcurrentDictionary<VBComponent, Task<Action>>();

        private readonly ConcurrentDictionary<VBComponent, Task> _resolverTasks =
            new ConcurrentDictionary<VBComponent, Task>();

        private readonly ConcurrentDictionary<VBComponent, CancellationTokenSource> _tokenSources =
            new ConcurrentDictionary<VBComponent, CancellationTokenSource>();

        private readonly Action<VBComponent, RubberduckParserState.State> _setParserState;

        public ParseCoordinator(Action<VBComponent, RubberduckParserState.State> setParserState)
        {
            _setParserState = setParserState;
        }

        public void Start(VBComponent component, Func<Action> parseAction)
        {
            var tokenSource = UpdateTokenSource(component);
            StartParserTask(component, parseAction, tokenSource);
        }

        private CancellationTokenSource UpdateTokenSource(VBComponent component)
        {
            CancellationTokenSource existingTokenSource;
            if (_tokenSources.TryGetValue(component, out existingTokenSource))
            {
                try
                {
                    existingTokenSource.Cancel();
                }
                catch (ObjectDisposedException)
                {
                }
            }

            var tokenSource = new CancellationTokenSource();
            _tokenSources[component] = tokenSource;
            return tokenSource;
        }

        private void StartParserTask(VBComponent component, Func<Action> parseAction, CancellationTokenSource tokenSource)
        {
            Task<Action> existingParseTask;
            if (_parseTasks.TryGetValue(component, out existingParseTask) && existingParseTask.Status == TaskStatus.Running)
            {
                // wait for the task to actually respond to cancellation
                existingParseTask.Wait();
            }

            _setParserState.Invoke(component, RubberduckParserState.State.Parsing);
            CancellationToken token;
            try
            {
                token = tokenSource.Token;
            }
            catch (ObjectDisposedException)
            {
                UpdateTokenSource(component);
                token = _tokenSources[component].Token;
            }

            _parseTasks[component] = Task.Factory.StartNew(parseAction, token);
            _parseTasks[component].ContinueWith(t =>
            {
                if (t.IsFaulted)
                {
                    _setParserState.Invoke(component, RubberduckParserState.State.Error);
                }
                else
                {
                    if (t.IsCompleted)
                    {
                        SetResolverTask(component, t.Result, token);
                        ResolveWhenReady(token);
                    }
                    else
                    {
                        Task<Action> parseTask;
                        _parseTasks.TryRemove(component, out parseTask);
                    }
                }
            }, token)
            .ContinueWith(t =>
            {
                CancellationTokenSource cts;
                _tokenSources.TryRemove(component, out cts);
                if (cts != null)
                {
                    cts.Dispose();
                }
            });
        }

        private void SetResolverTask(VBComponent component, Action resolverAction, CancellationToken token)
        {
            if (resolverAction == null)
            {
                return;
            }

            Task existingResolverTask;
            if (_resolverTasks.TryGetValue(component, out existingResolverTask) && existingResolverTask.Status == TaskStatus.Running)
            {
                // wait for the task to actually respond to cancellation
                existingResolverTask.Wait();
            }

            _resolverTasks[component] = new Task(resolverAction, token);
        }

        private void ResolveWhenReady(CancellationToken token)
        {
            var parseTasks = _parseTasks.Values.ToArray();
            Task.WaitAll(parseTasks, token);

            if (_parseTasks.Values.Any(task => !task.IsCompleted || task.IsCanceled))
            {
                return;
            }

            foreach (var resolverTask in _resolverTasks)
            {
                var component = resolverTask.Key;
                _setParserState.Invoke(component, RubberduckParserState.State.Resolving);
                resolverTask.Value.Start();
                resolverTask.Value.ContinueWith(t => ReportReadyState(component), token);
            }
        }

        private void ReportReadyState(VBComponent component)
        {
            _setParserState.Invoke(component, RubberduckParserState.State.Ready);
        }
    }
}
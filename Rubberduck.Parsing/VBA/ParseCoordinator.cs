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

        private readonly Action<RubberduckParserState.State> _setParserState;

        public ParseCoordinator(Action<RubberduckParserState.State> setParserState)
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
                existingTokenSource.Cancel();
            }

            var tokenSource = new CancellationTokenSource();
            _tokenSources[component] = tokenSource;
            return tokenSource;
        }

        private void StartParserTask(VBComponent component, Func<Action> parseAction, CancellationTokenSource tokenSource)
        {
            Task<Action> existingParseTask;
            if (_parseTasks.TryGetValue(component, out existingParseTask))
            {
                // wait for the task to actually respond to cancellation
                existingParseTask.Wait();
            }

            _setParserState.Invoke(RubberduckParserState.State.Parsing);

            _parseTasks[component] = Task.Factory
                .StartNew(parseAction, tokenSource.Token);

            _parseTasks[component].ContinueWith(t =>
            {
                if (t.IsFaulted)
                {
                    _setParserState.Invoke(RubberduckParserState.State.Error);
                }
                else
                {
                    SetResolverTask(component, t.Result, tokenSource.Token);
                    ResolveWhenReady(tokenSource.Token);
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
            if (_resolverTasks.TryGetValue(component, out existingResolverTask))
            {
                // wait for the task to actually respond to cancellation
                existingResolverTask.Wait();
            }

            _resolverTasks[component] = new Task(resolverAction, token)
                .ContinueWith(t => ReportReadyState(token), token);
        }

        private void ResolveWhenReady(CancellationToken token)
        {
            var parseTasks = _parseTasks.Values.ToArray();
            Task.WaitAll(parseTasks, token);

            if (_parseTasks.Values.Any(task => !task.IsCompleted || task.IsCanceled))
            {
                return;
            }

            _setParserState.Invoke(RubberduckParserState.State.Resolving);
            foreach (var resolverTask in _resolverTasks.Values)
            {
                resolverTask.Start();
            }
        }

        private void ReportReadyState(CancellationToken token)
        {
            var resolverTasks = _resolverTasks.Values.ToArray();
            Task.WaitAll(resolverTasks, token);

            if (_resolverTasks.Values.Any(task => !task.IsCompleted || task.IsCanceled))
            {
                return;
            }

            _setParserState.Invoke(RubberduckParserState.State.Ready);
        }
    }
}
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
        private readonly ConcurrentDictionary<VBComponent, Task> _parseTasks =
            new ConcurrentDictionary<VBComponent, Task>();

        private readonly ConcurrentDictionary<VBComponent, Task> _resolverTasks =
            new ConcurrentDictionary<VBComponent, Task>();

        private readonly Action<VBComponent, RubberduckParserState.State> _setParserState;

        public ParseCoordinator(Action<VBComponent, RubberduckParserState.State> setParserState)
        {
            _setParserState = setParserState;
        }

        public async Task StartAsync(VBComponent component, Task parseTask)
        {
            _setParserState.Invoke(component, RubberduckParserState.State.Parsing);
            _parseTasks[component] = parseTask;

            await parseTask.ContinueWith(t =>
            {
                if (t.IsFaulted)
                {
                    _setParserState.Invoke(component, RubberduckParserState.State.Error);
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
            if (_resolverTasks.TryGetValue(component, out existingResolverTask) 
                && existingResolverTask.Status == TaskStatus.Running
                && token.IsCancellationRequested)
            {
                // wait for the task to actually respond to cancellation
                existingResolverTask.Wait();
            }

            _resolverTasks[component] = new Task(resolverAction, token);
        }

        private void ResolveWhenReady(CancellationToken token)
        {
            if (!_parseTasks.Values.All(task => task.IsCompleted))
            {
                return;
            }

            foreach (var resolverTask in _resolverTasks.Where(t => t.Value.Status == TaskStatus.WaitingToRun))
            {
                var component = resolverTask.Key;
                _setParserState.Invoke(component, RubberduckParserState.State.Resolving);
                try
                {
                    resolverTask.Value.Start();
                    resolverTask.Value.ContinueWith(t => ReportReadyState(component), token);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }

        private void ReportReadyState(VBComponent component)
        {
            _setParserState.Invoke(component, RubberduckParserState.State.Ready);
        }
    }
}
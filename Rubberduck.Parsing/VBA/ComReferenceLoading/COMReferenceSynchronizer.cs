using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement;

namespace Rubberduck.Parsing.VBA.ComReferenceLoading
{
    public class COMReferenceSynchronizer : COMReferenceSynchronizerBase
    {
        private const int _maxReferenceLoadingConcurrency = -1;

        public COMReferenceSynchronizer(
            RubberduckParserState state, 
            IParserStateManager parserStateManager, 
            IProjectsProvider projectsProvider,
            IReferencedDeclarationsCollector referencedDeclarationsCollector)
        :base(
            state, 
            parserStateManager,
            projectsProvider,
            referencedDeclarationsCollector)
        { }


        protected override void LoadReferences(IEnumerable<ReferenceInfo> referencesToLoad, ConcurrentBag<ReferenceInfo> unmapped, CancellationToken token)
        {
            var referenceLoadingTaskScheduler = ThrottledTaskScheduler(_maxReferenceLoadingConcurrency);

            //Parallel.ForEach is not used because loading the references can contain IO-bound operations.
            var loadTasks = new List<Task>();
            foreach (var reference in referencesToLoad)
            {
                loadTasks.Add(Task.Factory.StartNew(
                                    () => LoadReference(reference, unmapped),
                                    token,
                                    TaskCreationOptions.None,
                                    referenceLoadingTaskScheduler
                                ));
            }

            try
            {
                Task.WaitAll(loadTasks.ToArray(), token);
            }
            catch (AggregateException exception)
            {
                if (exception.Flatten().InnerExceptions.All(ex => ex is OperationCanceledException))
                {
                    throw exception.InnerException ?? exception; //This eliminates the stack trace, but for the cancellation, this is irrelevant.
                }
                _parserStateManager.SetStatusAndFireStateChanged(this, ParserState.Error, token);
                throw;
            }
            token.ThrowIfCancellationRequested();
        }

        private TaskScheduler ThrottledTaskScheduler(int maxLevelOfConcurrency)
        {
            if (maxLevelOfConcurrency <= 0)
            {
                return TaskScheduler.Default;
            }
            else
            {
                var taskSchedulerPair = new ConcurrentExclusiveSchedulerPair(TaskScheduler.Default, maxLevelOfConcurrency);
                return taskSchedulerPair.ConcurrentScheduler;
            }
        }
    }
}

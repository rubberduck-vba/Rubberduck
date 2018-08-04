using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.UIContext;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;

namespace Rubberduck.UnitTesting
{
    public class TestEngine : ITestEngine
    {
        private readonly RubberduckParserState _state;
        private readonly IFakesFactory _fakesFactory;
        private readonly IVBETypeLibsAPI _typeLibApi;
        private readonly IUiDispatcher _uiDispatcher;

        public ParserState[] AllowedRunStates => new[]
        {
            ParserState.ResolvedDeclarations,
            ParserState.ResolvingReferences, 
            ParserState.Ready
        };

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private bool _testRequested;
        private readonly Dictionary<TestMethod, TestOutcome> testResults;
        public IEnumerable<TestMethod> Tests { get; }

        public TestOutcome RunAggregateOutcome {  get
            {
                if (testResults.Values.Any(o => o == TestOutcome.Failed))
                {
                    return TestOutcome.Failed;
                }
                if (testResults.Values.Any(o => o == TestOutcome.Inconclusive || o == TestOutcome.Ignored))
                {
                    return TestOutcome.Inconclusive;
                }
                if (testResults.Values.Any(o => o == TestOutcome.Succeeded))
                {
                    return TestOutcome.Succeeded;
                }
                // no test values recorded -> no tests run -> unknown
                return TestOutcome.Unknown;
            }
        }

        public TestEngine(RubberduckParserState state, IFakesFactory fakesFactory, IVBETypeLibsAPI typeLibApi, IUiDispatcher uiDispatcher)
        {
            Debug.WriteLine("TestEngine created.");
            _state = state;
            _fakesFactory = fakesFactory;
            _typeLibApi = typeLibApi;
            _uiDispatcher = uiDispatcher;
            
            _state.StateChanged += StateChangedHandler;
        }

        private void StateChangedHandler(object sender, ParserStateEventArgs e)
        {
            if (_testRequested && (e.State == ParserState.Ready))
            {
                _testRequested = false;
                _uiDispatcher.InvokeAsync(() =>
                {
                    RunInternal(Tests);
                });
            }

            if (_testRequested && !e.IsError)
            {
                _testRequested = false;
            }
        }

        public event EventHandler<TestCompletedEventArgs> TestCompleted;
        public event EventHandler TestsRefreshed;

        private void OnTestCompleted(TestMethod test, TestResult result)
        {
            var handler = TestCompleted;
            handler?.Invoke(this, new TestCompletedEventArgs(test, result));
        }

        public void Run()
        {
            _testRequested = true;
            Run(Tests);
        }

        public void Run(IEnumerable<TestMethod> tests)
        {
            _uiDispatcher.InvokeAsync(() => RunInternal(tests));
        }

        private void RunInternal(IEnumerable<TestMethod> tests)
        {
            if (!AllowedRunStates.Contains(_state.Status))
            {
                return;
            }

            _state.OnSuspendParser(this, AllowedRunStates, () => RunWhileSuspended(tests));
        }

        private void RunWhileSuspended(IEnumerable<TestMethod> tests)
        {
            var testMethods = tests as IList<TestMethod> ?? tests.ToList();
            if (!testMethods.Any())
            {
                return;
            }

            try
            {
                var modules = testMethods.GroupBy(test => test.Declaration.QualifiedName.QualifiedModuleName);
                foreach (var module in modules)
                {
                    var testInitialize = module.Key.FindTestInitializeMethods(_state).ToList();
                    var testCleanup = module.Key.FindTestCleanupMethods(_state).ToList();

                    var capturedModule = module;
                    var moduleTestMethods = testMethods
                        .Where(test =>
                        {
                            var qmn = test.Declaration.QualifiedName.QualifiedModuleName;

                            return qmn.ProjectId == capturedModule.Key.ProjectId
                                   && qmn.ComponentName == capturedModule.Key.ComponentName;
                        });

                    var fakes = _fakesFactory.Create();
                    var initializeMethods = module.Key.FindModuleInitializeMethods(_state);
                    try
                    {
                        RunInternal(initializeMethods);
                    }
                    catch (COMException ex)
                    {
                        Logger.Error(ex,
                            "Unexpected COM exception while initializing tests for module {0}. The module will be skipped.",
                            module.Key.Name);
                        foreach (var method in moduleTestMethods)
                        {
                            method.UpdateResult(TestOutcome.Unknown, AssertMessages.Assert_ComException);
                            OnTestCompleted(method, new TestResult(TestOutcome.Unknown));
                        }
                        continue;
                    }
                    foreach (var test in moduleTestMethods)
                    {
                        // no need to run setup/teardown for ignored tests
                        if (test.Declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.IgnoreTest))
                        {
                            test.UpdateResult(TestOutcome.Ignored);
                            OnTestCompleted(test, new TestResult(TestOutcome.Ignored));
                            continue;
                        }

                        var stopwatch = new Stopwatch();
                        stopwatch.Start();

                        try
                        {
                            fakes.StartTest();
                            RunInternal(testInitialize);
                            test.Run();
                            RunInternal(testCleanup);
                        }
                        catch (COMException ex)
                        {
                            Logger.Error(ex, "Unexpected COM exception while running tests.");
                            test.UpdateResult(TestOutcome.Inconclusive, AssertMessages.Assert_ComException);
                        }
                        finally
                        {
                            fakes.StopTest();
                        }

                        stopwatch.Stop();
                        test.Result.SetDuration(stopwatch.ElapsedMilliseconds);
                        OnTestCompleted(test, new TestResult(TestOutcome.Succeeded, duration: stopwatch.ElapsedMilliseconds));
                    }
                    var cleanupMethods = module.Key.FindModuleCleanupMethods(_state);
                    try
                    {
                        RunInternal(cleanupMethods);
                    }
                    catch (COMException ex)
                    {
                        Logger.Error(ex,
                            "Unexpected COM exception while cleaning up tests for module {0}. Aborting any further unit tests",
                            module.Key.Name);
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error(ex, "Unexpected expection while running unit tests; unit tests will be aborted");
            }
        }

        private void RunInternal(IEnumerable<Declaration> members)
        {
            var groupedMembers = members.GroupBy(m => m.ProjectId);
            foreach (var group in groupedMembers)
            {
                var project = _state.ProjectsProvider.Project(group.Key);
                using (var typeLib = TypeLibWrapper.FromVBProject(project))
                {
                    foreach (var member in group)
                    {
                        _typeLibApi.ExecuteCode(typeLib, member.QualifiedModuleName.ComponentName,
                            member.QualifiedName.MemberName);
                    }
                }
            }
        }
    }
}

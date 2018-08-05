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
    internal class TestEngine : ITestEngine
    {
        private static readonly ParserState[] AllowedRunStates = new ParserState[]
        {
            ParserState.ResolvedDeclarations,
            ParserState.ResolvingReferences, 
            ParserState.Ready
        };
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly RubberduckParserState _state;
        private readonly IFakesFactory _fakesFactory;
        private readonly IVBETypeLibsAPI _typeLibApi;
        private readonly IUiDispatcher _uiDispatcher;
        private bool _testRequested;
        private readonly Dictionary<TestMethod, TestOutcome> testResults = new Dictionary<TestMethod, TestOutcome>();
        public IEnumerable<TestMethod> Tests { get; private set; }

        public TestOutcome CurrentAggregateOutcome {  get
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
            // if we could run with the parser state change, tests should be updated
            if (CanRun())
            {
                Tests = TestDiscovery.GetAllTests(_state);
                TestsRefreshed?.Invoke(this, EventArgs.Empty);

                if (_testRequested)
                {
                    _testRequested = false;
                    _uiDispatcher.InvokeAsync(() =>
                    {
                        RunInternal(Tests);
                    });
                }
            }

            // any error cancels outstanding test runs
            if (e.IsError)
            {
                _testRequested = false;
            }
        }

        public event EventHandler<TestCompletedEventArgs> TestCompleted;
        public event EventHandler TestsRefreshed;

        private void OnTestCompleted(TestMethod test, TestResult result)
        {
            testResults.Add(test, result.Outcome);
            TestCompleted?.Invoke(this, new TestCompletedEventArgs(test, result));
        }

        public void RunAll()
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
            if (!CanRun())
            {
                return;
            }
            // FIXME we shouldn't need to handle awaiting a certain parser state ourselves anymore, right?
            // that would drop the _testsRequested member completely
            _state.OnSuspendParser(this, AllowedRunStates, () => RunWhileSuspended(tests));
        }

        private void RunWhileSuspended(IEnumerable<TestMethod> tests)
        {
            var testMethods = tests as IList<TestMethod> ?? tests.ToList();
            if (!testMethods.Any())
            {
                return;
            }
            testResults.Clear();
            try
            {
                var modules = testMethods.GroupBy(test => test.Declaration.QualifiedName.QualifiedModuleName);
                foreach (var module in modules)
                {
                    var testInitialize = TestDiscovery.FindTestInitializeMethods(module.Key, _state).ToList();
                    var testCleanup = TestDiscovery.FindTestCleanupMethods(module.Key, _state).ToList();

                    var capturedModule = module;
                    var moduleTestMethods = testMethods
                        .Where(test =>
                        {
                            var qmn = test.Declaration.QualifiedName.QualifiedModuleName;

                            return qmn.ProjectId == capturedModule.Key.ProjectId
                                   && qmn.ComponentName == capturedModule.Key.ComponentName;
                        });

                    var fakes = _fakesFactory.Create();
                    var initializeMethods = TestDiscovery.FindModuleInitializeMethods(module.Key, _state);
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
                            OnTestCompleted(method, new TestResult(TestOutcome.Unknown, AssertMessages.Assert_ComException));
                        }
                        continue;
                    }
                    foreach (var test in moduleTestMethods)
                    {
                        // no need to run setup/teardown for ignored tests
                        if (test.Declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.IgnoreTest))
                        {
                            OnTestCompleted(test, new TestResult(TestOutcome.Ignored));
                            continue;
                        }

                        try
                        {
                            fakes.StartTest();
                            try
                            {
                                RunInternal(testInitialize);
                            }
                            catch (Exception trace)
                            {
                                OnTestCompleted(test, new TestResult(TestOutcome.Inconclusive, AssertMessages.Assert_TestInitializeFailure));
                                Logger.Trace(trace, "Unexpected Exception when running TestInitialize");
                                continue;
                            }
                            var result = test.Run();
                            // we can trigger this event, because cleanup can fail without affecting the result
                            OnTestCompleted(test, result);
                            RunInternal(testCleanup);
                        }
                        catch (COMException ex)
                        {
                            Logger.Error(ex, "Unexpected COM exception while running tests.");
                            OnTestCompleted(test, new TestResult(TestOutcome.Inconclusive, AssertMessages.Assert_ComException));
                        }
                        finally
                        {
                            fakes.StopTest();
                        }
                    }
                    var cleanupMethods = TestDiscovery.FindModuleCleanupMethods(module.Key, _state);
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

        public bool CanRun()
        {
            return AllowedRunStates.Contains(_state.Status);
        }
    }
}

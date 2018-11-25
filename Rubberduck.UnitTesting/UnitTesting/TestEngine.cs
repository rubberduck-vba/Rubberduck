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
using Rubberduck.Parsing.VBA.Extensions;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting
{
    // FIXME litter logging around here
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
        private readonly IVBEInteraction _declarationRunner;
        private readonly ITypeLibWrapperProvider _wrapperProvider;
        private readonly IUiDispatcher _uiDispatcher;
        private readonly IVBE _vbe;

        private readonly List<TestMethod> LastRun = new List<TestMethod>();
        private readonly Dictionary<TestOutcome, List<TestMethod>> resultsByOutcome = new Dictionary<TestOutcome, List<TestMethod>>();
        public IEnumerable<TestMethod> Tests { get; private set; }
        public bool CanRun => AllowedRunStates.Contains(_state.Status) && _vbe.IsInDesignMode;
        public bool CanRepeatLastRun => LastRun.Any();
        
        private bool refreshBackoff;


        public TestEngine(RubberduckParserState state, IFakesFactory fakesFactory, IVBEInteraction declarationRunner, ITypeLibWrapperProvider wrapperProvider, IUiDispatcher uiDispatcher, IVBE vbe)
        {
            Debug.WriteLine("TestEngine created.");
            _state = state;
            _fakesFactory = fakesFactory;
            _declarationRunner = declarationRunner;
            _wrapperProvider = wrapperProvider;
            _uiDispatcher = uiDispatcher;
            _vbe = vbe;

            _state.StateChanged += StateChangedHandler;

            // avoid nulls in results by outcome
            foreach (TestOutcome outcome in Enum.GetValues(typeof(TestOutcome)))
            {
                resultsByOutcome.Add(outcome, new List<TestMethod>());
            }
        }

        public TestOutcome CurrentAggregateOutcome
        {
            get
            {
                if (resultsByOutcome[TestOutcome.Failed].Any())
                {
                    return TestOutcome.Failed;
                }
                if (resultsByOutcome[TestOutcome.Inconclusive].Any() || resultsByOutcome[TestOutcome.Ignored].Any())
                {
                    return TestOutcome.Inconclusive;
                }
                if (resultsByOutcome[TestOutcome.Succeeded].Any())
                {
                    return TestOutcome.Succeeded;
                }
                // no test values recorded -> no tests run -> unknown
                return TestOutcome.Unknown;
            }
        }

        private void StateChangedHandler(object sender, ParserStateEventArgs e)
        {
            if (!CanRun || e.IsError)
            {
                refreshBackoff = false;
            }
            // CanRun returned true already, only refresh tests if we're not backed off
            else if (!refreshBackoff && e.OldState != ParserState.Busy)
            {
                refreshBackoff = true;
                Tests = TestDiscovery.GetAllTests(_state);
                _uiDispatcher.InvokeAsync(() => TestsRefreshed?.Invoke(this, EventArgs.Empty));
            }
        }

        public event EventHandler<TestCompletedEventArgs> TestCompleted;
        public event EventHandler<TestRunCompletedEventArgs> TestRunCompleted;
        public event EventHandler TestsRefreshed;

        private void OnTestCompleted(TestMethod test, TestResult result)
        {
            resultsByOutcome[result.Outcome].Add(test);
            LastRun.Add(test);
            _uiDispatcher.InvokeAsync(() => TestCompleted?.Invoke(this, new TestCompletedEventArgs(test, result)));
        }

        public void Run(IEnumerable<TestMethod> tests)
        {
            _uiDispatcher.InvokeAsync(() => RunInternal(tests));
        }

        public void RunByOutcome(TestOutcome outcome)
        {
            Run(resultsByOutcome[outcome]);
        }

        public void RepeatLastRun()
        {
            Run(LastRun);
        }

        private void RunInternal(IEnumerable<TestMethod> tests)
        {
            if (!CanRun)
            {
                return;
            }
            _state.OnSuspendParser(this, AllowedRunStates, () => RunWhileSuspended(tests));
        }

        private void EnsureRubberduckIsReferencedForEarlyBoundTests()
        {
            var projectIdsOfMembersUsingAddInLibrary = _state.DeclarationFinder.AllUserDeclarations
                .Where(member => member.AsTypeName == "Rubberduck.PermissiveAssertClass"
                                 || member.AsTypeName == "Rubberduck.AssertClass")
                .Select(member => member.ProjectId)
                .ToHashSet();
            var projectsUsingAddInLibrary = _state.DeclarationFinder
                .UserDeclarations(DeclarationType.Project)
                .Where(declaration => projectIdsOfMembersUsingAddInLibrary.Contains(declaration.ProjectId))
                .Select(declaration => declaration.Project);

            foreach (var project in projectsUsingAddInLibrary)
            {
                _declarationRunner.EnsureProjectReferencesUnitTesting(project);
            }
        }

        private void RunWhileSuspended(IEnumerable<TestMethod> tests)
        {
            var testMethods = tests as IList<TestMethod> ?? tests.ToList();
            if (!testMethods.Any())
            {
                return;
            }
            LastRun.Clear();
            foreach (var resultAggregator in resultsByOutcome.Values)
            {
                resultAggregator.Clear();
            }
            try
            {
                EnsureRubberduckIsReferencedForEarlyBoundTests();
            }
            catch (InvalidOperationException e)
            {
                Logger.Warn(e);
                foreach (var test in testMethods)
                {
                    OnTestCompleted(test, new TestResult(TestOutcome.Failed, AssertMessages.Prerequisite_EarlyBindingReferenceMissing, 0));
                }
                return;
            }

            var overallTime = new Stopwatch();
            overallTime.Start();
            try
            {
                var testsByModule = testMethods.GroupBy(test => test.Declaration.QualifiedName.QualifiedModuleName)
                    .ToDictionary(grouping => grouping.Key, grouping => grouping.ToList());
                
                foreach (var moduleName in testsByModule.Keys)
                {
                    var testInitialize = TestDiscovery.FindTestInitializeMethods(moduleName, _state).ToList();
                    var testCleanup = TestDiscovery.FindTestCleanupMethods(moduleName, _state).ToList();
                    
                    var moduleTestMethods = testsByModule[moduleName];

                    var fakes = _fakesFactory.Create();
                    using (var typeLibWrapper = _wrapperProvider.TypeLibWrapperFromProject(moduleName.ProjectId))
                    {
                        try
                        {
                            _declarationRunner.RunDeclarations(typeLibWrapper, TestDiscovery.FindModuleInitializeMethods(moduleName, _state));
                        }
                        catch (COMException ex)
                        {
                            Logger.Error(ex, "Unexpected COM exception while initializing tests for module {0}. The module will be skipped.", moduleName.Name);
                            foreach (var method in moduleTestMethods)
                            {
                                OnTestCompleted(method, new TestResult(TestOutcome.Unknown, AssertMessages.TestRunner_ModuleInitializeFailure));
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
                                    _declarationRunner.RunDeclarations(typeLibWrapper, testInitialize);
                                }
                                catch (COMException trace)
                                {
                                    OnTestCompleted(test, new TestResult(TestOutcome.Inconclusive, AssertMessages.TestRunner_TestInitializeFailure));
                                    Logger.Trace(trace, "Unexpected COMException when running TestInitialize");
                                    continue;
                                }
                                var result = RunTestMethod(typeLibWrapper, test);
                                // we can trigger this event, because cleanup can fail without affecting the result
                                OnTestCompleted(test, result);
                                try
                                {
                                    _declarationRunner.RunDeclarations(typeLibWrapper, testCleanup);
                                }
                                catch (COMException cleanupFail)
                                {
                                    // Apparently the user doesn't need to know when test results for subsequent tests could be incorrect
                                    Logger.Trace(cleanupFail, "Unexpected COMException when running TestCleanup");
                                }
                            }
                            finally
                            {
                                fakes.StopTest();
                            }
                        }
                        try
                        {
                            _declarationRunner.RunDeclarations(typeLibWrapper, TestDiscovery.FindModuleCleanupMethods(moduleName, _state));
                        }
                        catch (COMException ex)
                        {
                            // FIXME somehow notify the user of this mess
                            Logger.Error(ex,
                                "Unexpected COM exception while cleaning up tests for module {0}. Aborting any further unit tests",
                                moduleName.Name);
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // FIXME somehow notify the user of this mess
                Logger.Error(ex, "Unexpected expection while running unit tests; unit tests will be aborted");
            }
            overallTime.Stop();
            TestRunCompleted?.Invoke(this, new TestRunCompletedEventArgs(overallTime.ElapsedMilliseconds));
        }
        
        private TestResult RunTestMethod(ITypeLibWrapper typeLib, TestMethod test)
        {
            long duration = 0;
            try
            {
                var assertResults = new List<AssertCompletedEventArgs>();
                _declarationRunner.RunTestMethod(typeLib, test, (s, e) => assertResults.Add(e), out duration);
                return EvaluateResults(assertResults, duration);
            }
            catch (COMException e)
            {
                Logger.Info(e, "Unexpected COM exception while running test method.");
                return new TestResult(TestOutcome.Inconclusive, AssertMessages.TestRunner_ComException, duration);
            }
            catch (Exception e)
            {
                Logger.Error(e, "Unexpected exceptino while running test method.");
                return new TestResult(TestOutcome.Inconclusive, AssertMessages.TestRunner_ExceptionDuringRun, duration);
            }
        }

        private TestResult EvaluateResults(IEnumerable<AssertCompletedEventArgs> assertResults, long duration)
        {
            var result = new AssertCompletedEventArgs(TestOutcome.Succeeded);

            if (assertResults.Any(assertion => assertion.Outcome != TestOutcome.Succeeded))
            {
                result = assertResults.First(assertion => assertion.Outcome != TestOutcome.Succeeded);
            }

            return new TestResult(result.Outcome, result.Message, duration);
        }
    }
}

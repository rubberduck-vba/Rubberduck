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
using Rubberduck.UI.UnitTesting;
using Rubberduck.VBEditor.ComManagement.TypeLibs;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting
{
    public class TestEngine : ITestEngine
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;
        private readonly IFakesFactory _fakesFactory;
        private readonly IVBETypeLibsAPI _typeLibApi;
        private readonly IUiDispatcher _uiDispatcher;

        public ParserState[] AllowedRunStates => new[]
        {
            //TODO: work out a clean way for communicating allowable states to OnSuspendParse
            //ParserState.ResolvedDeclarations, ParserState.ResolvingReferences, 
            ParserState.Ready
        };

        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private bool _testRequested;
        private IEnumerable<TestMethod> _tests;

        public TestEngine(TestExplorerModel model, IVBE vbe, RubberduckParserState state, IFakesFactory fakesFactory, IVBETypeLibsAPI typeLibApi, IUiDispatcher uiDispatcher)
        {
            Debug.WriteLine("TestEngine created.");
            Model = model;
            _vbe = vbe;
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
                    RunInternal(_tests);
                });
            }

            if (_testRequested && !e.IsError)
            {
                _testRequested = false;
            }
        }

        public TestExplorerModel Model { get; }

        public event EventHandler TestCompleted;

        private void OnTestCompleted()
        {
            var handler = TestCompleted;
            handler?.Invoke(this, EventArgs.Empty);
        }

        public void Refresh()
        {
            Model.Refresh();
        }

        public void Run()
        {
            _testRequested = true;
            _tests = Model.LastRun;
            // We will run the tests once parsing has completed
            Refresh();
        }

        public void Run(IEnumerable<TestMethod> tests)
        {
            _uiDispatcher.InvokeAsync(() => RunInternal(tests));
        }

        private void RunInternal(IEnumerable<TestMethod> tests)
        {
            if (_state.Status != ParserState.Ready)
            {
                return;
            }

            _state.OnSuspendParser(this, () =>
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
                                test.Declaration.QualifiedName.QualifiedModuleName.ProjectId ==
                                capturedModule.Key.ProjectId
                                && test.Declaration.QualifiedName.QualifiedModuleName.ComponentName ==
                                capturedModule.Key.ComponentName);

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
                            }
                            continue;
                        }
                        foreach (var test in moduleTestMethods)
                        {
                            // no need to run setup/teardown for ignored tests
                            if (test.Declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.IgnoreTest))
                            {
                                test.UpdateResult(TestOutcome.Ignored);
                                OnTestCompleted();
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

                            OnTestCompleted();
                            Model.AddExecutedTest(test);
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
            });
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

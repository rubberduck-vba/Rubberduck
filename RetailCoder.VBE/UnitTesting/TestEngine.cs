﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI.UnitTesting;
using Rubberduck.VBEditor.Application;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UnitTesting
{
    public class TestEngine : ITestEngine
    {
        private readonly TestExplorerModel _model;
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;

        // can't be assigned from constructor because ActiveVBProject is null at startup:
        private IHostApplication _hostApplication; 

        public TestEngine(TestExplorerModel model, IVBE vbe, RubberduckParserState state)
        {
            _model = model;
            _vbe = vbe;
            _state = state;
        }

        public TestExplorerModel Model { get { return _model; } }

        public event EventHandler TestCompleted;

        private void OnTestCompleted()
        {
            var handler = TestCompleted;
            if (handler != null)
            {
                handler.Invoke(this, EventArgs.Empty);
            }
        }

        public void Refresh()
        {
            _model.Refresh();
        }

        public void Run()
        {
            Refresh();
            Run(_model.LastRun);
        }

        public void Run(IEnumerable<TestMethod> tests)
        {
            var testMethods = tests as IList<TestMethod> ?? tests.ToList();
            if (!testMethods.Any())
            {
                return;
            }

            var modules = testMethods.GroupBy(test => test.Declaration.QualifiedName.QualifiedModuleName);
            foreach (var module in modules)
            {
                var testInitialize = module.Key.FindTestInitializeMethods(_state).ToList();
                var testCleanup = module.Key.FindTestCleanupMethods(_state).ToList();

                var capturedModule = module;
                var moduleTestMethods = testMethods
                    .Where(test => test.Declaration.QualifiedName.QualifiedModuleName.ProjectId == capturedModule.Key.ProjectId
                                && test.Declaration.QualifiedName.QualifiedModuleName.ComponentName == capturedModule.Key.ComponentName);

                Run(module.Key.FindModuleInitializeMethods(_state));
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

                    Run(testInitialize);
                    test.Run();
                    Run(testCleanup);

                    stopwatch.Stop();
                    test.Result.SetDuration(stopwatch.ElapsedMilliseconds);

                    OnTestCompleted();
                    _model.AddExecutedTest(test);
                }
                Run(module.Key.FindModuleCleanupMethods(_state));
            }
        }

        private void Run(IEnumerable<Declaration> members)
        {
            if (_hostApplication == null)
            {
                _hostApplication = _vbe.HostApplication();
            }

            foreach (var member in members)
            {
                _hostApplication.Run(member);
            }
        }
    }
}

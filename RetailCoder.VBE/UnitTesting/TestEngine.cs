using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Reflection;
using Rubberduck.UI.UnitTesting;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.UnitTesting
{
    public class TestEngine : ITestEngine
    {
        private readonly TestExplorerModelBase _model;
        private readonly VBE _vbe;

        // can't be assigned from constructor because ActiveVBProject is null at startup:
        private IHostApplication _hostApplication; 

        public TestEngine(TestExplorerModelBase model, VBE vbe)
        {
            _model = model;
            _vbe = vbe;
        }

        public TestExplorerModelBase Model { get { return _model; } }

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

            var modules = testMethods.GroupBy(test => test.QualifiedMemberName.QualifiedModuleName);
            foreach (var module in modules)
            {
                var testInitialize = FindTestInitializeMethods(module.Key).ToList();
                var testCleanup = FindTestCleanupMethods(module.Key).ToList();

                Run(FindModuleInitializeMethods(module.Key));
                foreach (var test in module)
                {
                    Run(testInitialize);
                    test.Run();

                    OnTestCompleted();
                    _model.AddExecutedTest(test);

                    Run(testCleanup);

                }
                Run(FindModuleCleanupMethods(module.Key));
            }
        }

        private void Run(IEnumerable<QualifiedMemberName> members)
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

        private static IEnumerable<QualifiedMemberName> FindModuleInitializeMethods(QualifiedModuleName module)
        {
            return module.Component.GetMembers(vbext_ProcKind.vbext_pk_Proc)
                .Where(m => m.HasAttribute<ModuleInitializeAttribute>())
                .Select(m => m.QualifiedMemberName);
        }

        private static IEnumerable<QualifiedMemberName> FindModuleCleanupMethods(QualifiedModuleName module)
        {
            return module.Component.GetMembers(vbext_ProcKind.vbext_pk_Proc)
                .Where(m => m.HasAttribute<ModuleCleanupAttribute>())
                .Select(m => m.QualifiedMemberName);
        }

        private static IEnumerable<QualifiedMemberName> FindTestInitializeMethods(QualifiedModuleName module)
        {
            return module.Component.GetMembers(vbext_ProcKind.vbext_pk_Proc)
                .Where(m => m.HasAttribute<TestInitializeAttribute>())
                .Select(m => m.QualifiedMemberName);
        }

        private static IEnumerable<QualifiedMemberName> FindTestCleanupMethods(QualifiedModuleName module)
        {
            return module.Component.GetMembers(vbext_ProcKind.vbext_pk_Proc)
                .Where(m => m.HasAttribute<TestCleanupAttribute>())
                .Select(m => m.QualifiedMemberName);
        }
    }
}

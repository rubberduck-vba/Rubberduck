using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.UI;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.UnitTesting
{
    public class TestMethod : ViewModelBase, IEquatable<TestMethod>, IEditableObject, INavigateSource
    {
        private readonly ICollection<TestResult> _assertResults = new List<TestResult>();
        private readonly IHostApplication _hostApp;

        public TestMethod(QualifiedMemberName qualifiedMemberName, VBE vbe)
        {
            _qualifiedMemberName = qualifiedMemberName;
            _vbe = vbe;
            _hostApp = vbe.HostApplication();
        }

        private readonly QualifiedMemberName _qualifiedMemberName;
        private readonly VBE _vbe;
        public QualifiedMemberName QualifiedMemberName { get { return _qualifiedMemberName; } }

        public void Run()
        {
            _assertResults.Clear(); //clear previous results to account for changes being made

            TestResult result;
            var duration = new TimeSpan();
            try
            {
                AssertHandler.OnAssertCompleted += HandleAssertCompleted;
                duration = _hostApp.TimedMethodCall(_qualifiedMemberName);
                AssertHandler.OnAssertCompleted -= HandleAssertCompleted;
                
                result = EvaluateResults();
            }
            catch(Exception exception)
            {
                result = TestResult.Inconclusive("Test raised an error. " + exception.Message);
            }
            
            Result = new TestResult(result, duration.Milliseconds);
        }

        private TestResult _result = TestResult.Unknown();
        public TestResult Result
        {
            get { return _result; } 
            set { _result = value; OnPropertyChanged();}
        }

        void HandleAssertCompleted(object sender, AssertCompletedEventArgs e)
        {
            _assertResults.Add(e.Result);
        }

        private TestResult EvaluateResults()
        {
            var result = TestResult.Success();

            if (_assertResults.Any(assertion => assertion.Outcome == TestOutcome.Failed || assertion.Outcome == TestOutcome.Inconclusive))
            {
                result = _assertResults.First(assertion => assertion.Outcome == TestOutcome.Failed || assertion.Outcome == TestOutcome.Inconclusive);
            }

            return result;
        }

        public NavigateCodeEventArgs GetNavigationArgs()
        {
            try
            {
                var moduleName = QualifiedMemberName.QualifiedModuleName;
                var methodName = QualifiedMemberName.MemberName;

                var module = _vbe.VBProjects.Cast<VBProject>()
                    .Single(project => project == QualifiedMemberName.QualifiedModuleName.Project)
                    .VBComponents.Cast<VBComponent>()
                    .Single(component => component.Name == QualifiedMemberName.QualifiedModuleName.ComponentName)
                    .CodeModule;

                var startLine = module.get_ProcStartLine(methodName, vbext_ProcKind.vbext_pk_Proc);
                var endLine = startLine + module.get_ProcCountLines(methodName, vbext_ProcKind.vbext_pk_Proc);
                var endLineColumns = module.get_Lines(endLine, 1).Length;

                var selection = new Selection(startLine, 1, endLine, endLineColumns == 0 ? 1 : endLineColumns);
                return new NavigateCodeEventArgs(new QualifiedSelection(moduleName, selection));
            }
            catch (COMException)
            {
                return null;
            }
        }

        public bool Equals(TestMethod other)
        {
            return QualifiedMemberName.Equals(other.QualifiedMemberName);
        }

        public override bool Equals(object obj)
        {
            return obj is TestMethod
                && ((TestMethod)obj).QualifiedMemberName.Equals(QualifiedMemberName);
        }

        public override int GetHashCode()
        {
            return QualifiedMemberName.GetHashCode();
        }

        private TestResult _cachedResult;

        private bool _isEditing;
        public bool IsEditing { get { return _isEditing; } set { _isEditing = value; OnPropertyChanged(); } }

        public void BeginEdit()
        {
            _cachedResult = new TestResult(Result, Result.Duration);
            IsEditing = true;
        }

        public void EndEdit()
        {
            _cachedResult = null;
            IsEditing = false;
        }

        public void CancelEdit()
        {
            if (_cachedResult != null)
            {
                Result = _cachedResult;
            }

            _cachedResult = null;
            IsEditing = false;
        }

        public override string ToString()
        {
            return string.Format("{0}: {1} ({2}ms) {3}", QualifiedMemberName, Result.Outcome, Result.Duration, Result.Output);
        }
    }
}

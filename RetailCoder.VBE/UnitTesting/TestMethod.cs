using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Microsoft.Vbe.Interop;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;
using Rubberduck.UI.Controls;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.Extensions;
using Rubberduck.VBEditor.VBEHost;

namespace Rubberduck.UnitTesting
{
    public class TestMethod : ViewModelBase, IEquatable<TestMethod>, INavigateSource
    {
        private readonly ICollection<AssertCompletedEventArgs> _assertResults = new List<AssertCompletedEventArgs>();
        private readonly IHostApplication _hostApp;

        public TestMethod(Declaration declaration, VBE vbe)
        {
            _declaration = declaration;
            _hostApp = vbe.HostApplication();
        }

        private Declaration _declaration;
        public Declaration Declaration { get { return _declaration; } }

        public void SetDeclaration(Declaration declaration)
        {
            _declaration = declaration;
            OnPropertyChanged("Declaration");
        }

        public void Run()
        {
            _assertResults.Clear(); //clear previous results to account for changes being made

            AssertCompletedEventArgs result;
            var duration = new TimeSpan();
            var startTime = DateTime.Now;
            try
            {
                AssertHandler.OnAssertCompleted += HandleAssertCompleted;
                _hostApp.Run(Declaration.QualifiedName);
                AssertHandler.OnAssertCompleted -= HandleAssertCompleted;
                
                result = EvaluateResults();
            }
            catch(Exception exception)
            {
                result = new AssertCompletedEventArgs(TestOutcome.Inconclusive, "Test raised an error. " + exception.Message);
            }
            var endTime = DateTime.Now;
            UpdateResult(result.Outcome, result.Message, duration.Milliseconds, startTime, endTime);
        }
        
        public void UpdateResult(TestOutcome outcome, string message = "", long duration = 0, DateTime? startTime = null, DateTime? endTime = null)
        {
            Result.SetValues(outcome, message, duration, startTime, endTime);
            OnPropertyChanged("Result");
        }

        private TestResult _result = new TestResult(TestOutcome.Unknown);
        public TestResult Result
        {
            get { return _result; } 
            set { _result = value; OnPropertyChanged(); }
        }

        void HandleAssertCompleted(object sender, AssertCompletedEventArgs e)
        {
            _assertResults.Add(e);
        }

        private AssertCompletedEventArgs EvaluateResults()
        {
            var result = new AssertCompletedEventArgs(TestOutcome.Succeeded);

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
                var moduleName = Declaration.QualifiedName.QualifiedModuleName;
                var methodName = Declaration.IdentifierName;
                var module = moduleName.Component.CodeModule;

                var startLine = module.ProcStartLine[methodName, vbext_ProcKind.vbext_pk_Proc];
                var endLine = startLine + module.ProcCountLines[methodName, vbext_ProcKind.vbext_pk_Proc];
                var endLineColumns = module.Lines[endLine, 1].Length;

                var selection = new Selection(startLine, 1, endLine, endLineColumns == 0 ? 1 : endLineColumns);
                return new NavigateCodeEventArgs(new QualifiedSelection(moduleName, selection));
            }
            catch (COMException)
            {
                return null;
            }
        }

        public object[] ToArray()
        {
            return new object[] { Declaration.QualifiedName.QualifiedModuleName.ProjectTitle, Declaration.QualifiedName.QualifiedModuleName.ComponentTitle, Declaration.IdentifierName, 
                _result.Outcome.ToString(), _result.Output, _result.StartTime.ToString(CultureInfo.InvariantCulture), _result.EndTime.ToString(CultureInfo.InvariantCulture), _result.Duration };
        }

        public bool Equals(TestMethod other)
        {
            return Declaration.QualifiedName.Equals(other.Declaration.QualifiedName);
        }

        public override bool Equals(object obj)
        {
            return obj is TestMethod
                && ((TestMethod)obj).Declaration.QualifiedName.Equals(Declaration.QualifiedName);
        }

        public override int GetHashCode()
        {
            return Declaration.QualifiedName.GetHashCode();
        }

        public override string ToString()
        {
            return string.Format("{0}: {1} ({2}ms) {3}", Declaration.QualifiedName, Result.Outcome, Result.Duration, Result.Output);
        }
    }
}

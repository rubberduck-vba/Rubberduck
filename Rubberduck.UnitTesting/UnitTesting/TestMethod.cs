using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Interaction.Navigation;
using System.Diagnostics;

namespace Rubberduck.UnitTesting
{
    [SuppressMessage("ReSharper", "ExplicitCallerInfoArgument")]
    public class TestMethod : IEquatable<TestMethod>, INavigateSource
    {
        private readonly ICollection<AssertCompletedEventArgs> _assertResults = new List<AssertCompletedEventArgs>();
        private readonly IVBE _vbe;
        private readonly IVBETypeLibsAPI _typeLibApi;
        private readonly RubberduckParserState _state;

        public TestMethod(RubberduckParserState state, Declaration declaration, IVBETypeLibsAPI typeLibApi)
        {
            _state = state;
            Declaration = declaration;
            _typeLibApi = typeLibApi;
        }
        public Declaration Declaration { get; }

        public TestResult Run()
        {
            _assertResults.Clear(); //clear previous results to account for changes being made

            AssertCompletedEventArgs result;
            var duration = new Stopwatch();
            try
            {
                AssertHandler.OnAssertCompleted += HandleAssertCompleted;
                var project = _state.ProjectsProvider.Project(Declaration.ProjectId);

                duration.Start();

                _typeLibApi.ExecuteCode(project, Declaration.QualifiedModuleName.ComponentName,
                    Declaration.QualifiedName.MemberName);

                duration.Stop();
                AssertHandler.OnAssertCompleted -= HandleAssertCompleted;
                result = EvaluateResults();
            }
            catch(Exception exception)
            {
                result = new AssertCompletedEventArgs(TestOutcome.Inconclusive, "Test raised an error. " + exception.Message);
            }
            return new TestResult(result.Outcome, result.Message, duration.ElapsedMilliseconds);
        }
        
        public void UpdateResult(TestOutcome outcome, string message = "", long duration = 0, DateTime? startTime = null, DateTime? endTime = null)
        {
            Result.SetValues(outcome, message, duration, startTime, endTime);
            //OnPropertyChanged("Result");
        }

        private TestResult _result = new TestResult(TestOutcome.Unknown);
        public TestResult Result
        {
            get =>_result;
            set { _result = value; /*OnPropertyChanged();*/ }
        }

        public TestCategory Category
        {
            get
            {
                var testMethodAnnotation = (TestMethodAnnotation) Declaration.Annotations
                    .First(annotation => annotation.AnnotationType == AnnotationType.TestMethod);

                return new TestCategory(testMethodAnnotation.Category);
            }
        }

        private void HandleAssertCompleted(object sender, AssertCompletedEventArgs e)
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
            return new NavigateCodeEventArgs(new QualifiedSelection(Declaration.QualifiedName.QualifiedModuleName, Declaration.Context.GetSelection()));
        }

        public bool Equals(TestMethod other)
        {
            return other != null && Declaration.QualifiedName.Equals(other.Declaration.QualifiedName);
        }

        public override bool Equals(object obj)
        {
            return obj is TestMethod method && method.Declaration.QualifiedName.Equals(Declaration.QualifiedName);
        }

        public override int GetHashCode()
        {
            return Declaration.QualifiedName.GetHashCode();
        }
    }
}

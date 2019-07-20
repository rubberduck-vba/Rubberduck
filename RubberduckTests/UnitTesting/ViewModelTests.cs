using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Data;
using NUnit.Framework;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UI.UnitTesting.ViewModels;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    // FIXME - These commented tests need to be restored after TestEngine refactor.
    [NonParallelizable]
    [TestFixture, Apartment(ApartmentState.STA)]
    public class ViewModelTests
    {
        [Test]
        [TestCase(1)]
        [TestCase(2)]
        [TestCase(3)]
        [Category("Unit Testing")]
        public void UiDiscoversAnnotatedTestMethods(int testCount)
        {
            var engine = new MockedTestEngine(testCount);
            var model = new MockedTestExplorerModel(engine);
            using (var viewModel = new MockedTestExplorer(model))
            {
                engine.ParserState.OnParseRequested(engine);
                Assert.AreEqual(testCount, viewModel.ViewModel.Tests.OfType<TestMethodViewModel>().Count());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void UiRemovesRemovedTestMethods()
        {
            var engine = new MockedTestEngine(new List<string> { "TestModule1", "TestModule2" }, new List<int> { 1, 1 });
            var model = new MockedTestExplorerModel(engine);
            using (var viewModel = new MockedTestExplorer(model))
            {
                engine.ParserState.OnParseRequested(engine);
                Assert.AreEqual(2, viewModel.ViewModel.Tests.OfType<TestMethodViewModel>().Count());

                var project = engine.Vbe.Object.VBProjects.First();
                var component = project.VBComponents.First();
                project.VBComponents.Remove(component);

                engine.ParserState.OnParseRequested(engine);
                Assert.AreEqual(1, viewModel.ViewModel.Tests.OfType<TestMethodViewModel>().Count());
            }
        }

        [Test]
        [TestCase(TestExplorerGrouping.Outcome, "Result.Outcome")]
        [TestCase(TestExplorerGrouping.Location, "QualifiedName.QualifiedModuleName.Name")]
        [TestCase(TestExplorerGrouping.Category, "Method.Category.Name")]
        [Category("Unit Testing")]
        public void TestGrouping_ChangesUpdateGroups(TestExplorerGrouping grouping, string expected)
        {
            var engine = new MockedTestEngine(3);
            var model = new MockedTestExplorerModel(engine);
            using (var viewModel = new MockedTestExplorer(model))
            {
                engine.ParserState.OnParseRequested(engine);
                viewModel.ViewModel.TestGrouping = grouping;

                var actual = ((PropertyGroupDescription)viewModel.ViewModel.Tests.GroupDescriptions.First()).PropertyName;

                Assert.AreEqual(1, viewModel.ViewModel.Tests.Groups.Count);
                Assert.AreEqual(expected, actual);
            }
        }

        private static readonly Dictionary<TestOutcome, (TestOutcome Outcome, string Output, long Duration)> DummyOutcomes = new Dictionary<TestOutcome, (TestOutcome, string, long)>
        {
            { TestOutcome.Succeeded,  (TestOutcome.Succeeded, "", 0)  },
            { TestOutcome.Inconclusive,  (TestOutcome.Inconclusive, "", 0)  },
            { TestOutcome.Failed,  (TestOutcome.Failed, "", 0)  },
            { TestOutcome.Ignored,  (TestOutcome.Ignored, "", 0)  }
        };


        //[Test]
        //[NonParallelizable]
        //[TestCase(new[] { TestOutcome.Succeeded, TestOutcome.Failed })]
        //[TestCase(new[] { TestOutcome.Succeeded, TestOutcome.Succeeded, TestOutcome.Succeeded })]
        //[TestCase(new[] { TestOutcome.Succeeded, TestOutcome.Inconclusive, TestOutcome.Failed })]
        //[TestCase(new[] { TestOutcome.Inconclusive, TestOutcome.Inconclusive, TestOutcome.Succeeded })]
        //[TestCase(new[] { TestOutcome.Failed, TestOutcome.Failed, TestOutcome.Failed })]
        //[TestCase(new[] { TestOutcome.Succeeded, TestOutcome.Ignored })]
        //[TestCase(new[] { TestOutcome.Succeeded, TestOutcome.Ignored, TestOutcome.Failed })]
        //public void TestGrouping_GroupsByOutcome(params TestOutcome[] tests)
        //{
        //    var underTest = tests.Select(test => DummyOutcomes[test]).ToList();
        //    var model = new MockedTestExplorerModel(underTest);

        //    using (var viewModel = new MockedTestExplorer(model))
        //    {
        //        viewModel.ViewModel.TestGrouping = TestExplorerGrouping.Outcome;

        //        model.Engine.ParserState.OnParseRequested(model);
        //        model.Model.ExecuteTests(model.Model.Tests);
        //        Thread.SpinWait(25);

        //        var actual = viewModel.ViewModel.Tests.Groups.Count;
        //        var expected = tests.Distinct().Count();

        //        Assert.AreEqual(expected, actual);
        //    }
        //}

        [Test]
        [Category("Unit Testing")]
        public void TestGrouping_GroupsByLocation()
        {
            var engine = new MockedTestEngine(new List<string> { "TestModule1", "TestModule2" }, new List<int> { 1, 1 });
            var model = new MockedTestExplorerModel(engine);
            using (var viewModel = new MockedTestExplorer(model))
            {
                viewModel.ViewModel.TestGrouping = TestExplorerGrouping.Location;
                engine.ParserState.OnParseRequested(engine);

                Assert.AreEqual(2, viewModel.ViewModel.Tests.Groups.Count);
            }
        }

        [Test]
        [TestCase("Foo", null, null)]
        [TestCase(null, null)]
        [TestCase("Foo", "Bar")]
        [TestCase("Foo", "Bar", "Foo", "Bar")]
        [TestCase("Foo", "Bar", "Baz")]
        [TestCase("Foo", "Bar", "Bar", "Baz")]
        [Category("Unit Testing")]
        public void TestGrouping_GroupsByCategory(params string[] categories)
        {
            var code = string.Join(Environment.NewLine,
                           Enumerable.Range(1, categories.Length)
                               .Select(num => MockedTestEngine.GetTestMethod(num, false, categories[num - 1])));

            var engine = new MockedTestEngine(code);
            var model = new MockedTestExplorerModel(engine);
            using (var viewModel = new MockedTestExplorer(model))
            {
                viewModel.ViewModel.TestGrouping = TestExplorerGrouping.Category;
                engine.ParserState.OnParseRequested(engine);

                var actual = viewModel.ViewModel.Tests.Groups.Count;
                var expected = categories.Distinct().Count();

                Assert.AreEqual(expected, actual);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void RunSingleTestCommand_DisabledNoSelectedTest()
        {
            var engine = new MockedTestEngine(3);
            var model = new MockedTestExplorerModel(engine);
            using (var viewModel = new MockedTestExplorer(model))
            {
                engine.ParserState.OnParseRequested(engine);
                Assert.IsFalse(viewModel.ViewModel.RunSingleTestCommand.CanExecute(null));
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void RunSingleTestCommand_RunsSelectedTest()
        {
            var engine = new MockedTestEngine(3);
            var model = new MockedTestExplorerModel(engine);
            using (var viewModel = new MockedTestExplorer(model))
            {
                engine.ParserState.OnParseRequested(engine);
                viewModel.ViewModel.MouseOverTest = model.Model.Tests.First();
                viewModel.ViewModel.RunSingleTestCommand.Execute(null);
                Assert.AreEqual(1, engine.TestEngine.LastRunTests.Count);
            }
        }

        [Test]
        [TestCase(0, false)]
        [TestCase(1, true)]
        [TestCase(2, true)]
        [TestCase(3, true)]
        [Category("Unit Testing")]
        public void RunSelectedTestsCommand_CanExecuteMultipleTests(int testCount, bool expected)
        {
            var engine = new MockedTestEngine(3);
            var model = new MockedTestExplorerModel(engine);
            using (var viewModel = new MockedTestExplorer(model))
            {
                engine.ParserState.OnParseRequested(engine);
                var tests = model.Model.Tests.Take(testCount).ToList();
                Assert.AreEqual(expected, viewModel.ViewModel.RunSelectedTestsCommand.CanExecute(tests));
            }
        }

        [Test]
        [TestCase(0)]
        [TestCase(1)]
        [TestCase(2)]
        [TestCase(3)]
        [Category("Unit Testing")]
        public void RunSelectedTestsCommand_ExecutesMultipleTests(int testCount)
        {
            var engine = new MockedTestEngine(3);
            var model = new MockedTestExplorerModel(engine);
            using (var viewModel = new MockedTestExplorer(model))
            {
                engine.ParserState.OnParseRequested(engine);
                var tests = model.Model.Tests.Take(testCount).ToList();
                viewModel.ViewModel.RunSelectedTestsCommand.Execute(tests);

                Assert.AreEqual(testCount, engine.TestEngine.LastRunTests.Count);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void CollapseAllCommand_SetsExpandedState()
        {
            var engine = new MockedTestEngine(1);
            var model = new MockedTestExplorerModel(engine);
            using (var viewModel = new MockedTestExplorer(model))
            {
                viewModel.ViewModel.ExpandedState = true;
                viewModel.ViewModel.CollapseAllCommand.Execute(null);

                Assert.IsFalse(viewModel.ViewModel.ExpandedState);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void ExpandAllCommand_SetsExpandedState()
        {
            var engine = new MockedTestEngine(1);
            var model = new MockedTestExplorerModel(engine);
            using (var viewModel = new MockedTestExplorer(model))
            {
                viewModel.ViewModel.ExpandedState = false;
                viewModel.ViewModel.ExpandAllCommand.Execute(null);

                Assert.IsTrue(viewModel.ViewModel.ExpandedState);
            }
        }
    }
}
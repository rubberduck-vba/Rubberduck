using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Windows.Media;
using Moq;
using NUnit.Framework;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    // FIXME - These commented tests need to be restored after TestEngine refactor.
    [NonParallelizable]
    [TestFixture, Apartment(ApartmentState.STA)]
    public class TestExplorerModelTests
    {
        [Test]
        [TestCase(1)]
        [TestCase(2)]
        [TestCase(3)]
        [Category("Unit Testing")]
        public void CurrentRunTestCount_MatchesEngine(int testCount)
        {
            var engine = new MockedTestEngine(testCount);
            using (var model = new MockedTestExplorerModel(engine))
            {
                engine.ParserState.OnParseRequested(engine);
                model.Model.ExecuteTests(model.Model.Tests);
                Assert.AreEqual(testCount, model.Model.CurrentRunTestCount);
            }
        }

        [Test]
        [TestCase(1)]
        [TestCase(2)]
        [TestCase(3)]
        [Category("Unit Testing")]
        public void ExecutedCount_MatchesNumberOfTestsRun(int testCount)
        {
            var engine = new MockedTestEngine(3);
            using (var model = new MockedTestExplorerModel(engine))
            {
                engine.ParserState.OnParseRequested(engine);
                model.Model.ExecuteTests(model.Model.Tests.Take(testCount).ToList());
                Assert.AreEqual(testCount, model.Model.ExecutedCount);
            }
        }

        private const int DummyTestDuration = 10;

        private static readonly Dictionary<TestOutcome, (TestOutcome Outcome, string Output, long Duration)> DummyOutcomes = new Dictionary<TestOutcome, (TestOutcome, string, long)>
        {
            { TestOutcome.Succeeded,  (TestOutcome.Succeeded, "", DummyTestDuration)  },
            { TestOutcome.Inconclusive,  (TestOutcome.Inconclusive, "", DummyTestDuration)  },
            { TestOutcome.Failed,  (TestOutcome.Failed, "", DummyTestDuration)  },
            { TestOutcome.Ignored,  (TestOutcome.Ignored, "", DummyTestDuration)  }
        };

        //[Test]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Succeeded, TestOutcome.Succeeded })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Inconclusive, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Inconclusive, TestOutcome.Inconclusive, TestOutcome.Inconclusive })]
        //[TestCase(new object[] { TestOutcome.Failed, TestOutcome.Failed, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Ignored })]
        //[TestCase(new object[] { TestOutcome.Ignored, TestOutcome.Ignored, TestOutcome.Ignored })]
        //[Category("Unit Testing")]
        //public void LastTestSucceededCount_CountIsCorrect(params TestOutcome[] tests)
        //{
        //    var underTest = tests.Select(test => DummyOutcomes[test]).ToList();

        //    using (var model = new MockedTestExplorerModel(underTest))
        //    {
        //        model.Engine.ParserState.OnParseRequested(model);
        //        model.Model.ExecuteTests(model.Model.Tests);
        //        Thread.SpinWait(25);

        //        var expected = tests.Count(outcome => outcome == TestOutcome.Succeeded);
        //        Assert.AreEqual(expected, model.Model.LastTestSucceededCount);
        //    }
        //}

        //[Test]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Succeeded, TestOutcome.Succeeded })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Inconclusive, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Inconclusive, TestOutcome.Inconclusive, TestOutcome.Inconclusive })]
        //[TestCase(new object[] { TestOutcome.Failed, TestOutcome.Failed, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Ignored })]
        //[TestCase(new object[] { TestOutcome.Ignored, TestOutcome.Ignored, TestOutcome.Ignored })]
        //[Category("Unit Testing")]
        //public void LastTestIgnoredCount_CountIsCorrect(params TestOutcome[] tests)
        //{
        //    var underTest = tests.Select(test => DummyOutcomes[test]).ToList();

        //    using (var model = new MockedTestExplorerModel(underTest))
        //    {
        //        model.Engine.ParserState.OnParseRequested(model);
        //        model.Model.ExecuteTests(model.Model.Tests);
        //        Thread.SpinWait(25);

        //        var expected = tests.Count(outcome => outcome == TestOutcome.Ignored);
        //        Assert.AreEqual(expected, model.Model.LastTestIgnoredCount);
        //    }
        //}

        //[Test]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Succeeded, TestOutcome.Succeeded })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Inconclusive, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Inconclusive, TestOutcome.Inconclusive, TestOutcome.Inconclusive })]
        //[TestCase(new object[] { TestOutcome.Failed, TestOutcome.Failed, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Ignored })]
        //[TestCase(new object[] { TestOutcome.Ignored, TestOutcome.Ignored, TestOutcome.Ignored })]
        //[Category("Unit Testing")]
        //public void LastTestInconclusiveCount_CountIsCorrect(params TestOutcome[] tests)
        //{
        //    var underTest = tests.Select(test => DummyOutcomes[test]).ToList();

        //    using (var model = new MockedTestExplorerModel(underTest))
        //    {
        //        model.Engine.ParserState.OnParseRequested(model);
        //        model.Model.ExecuteTests(model.Model.Tests);
        //        Thread.SpinWait(25);

        //        var expected = tests.Count(outcome => outcome == TestOutcome.Inconclusive);
        //        Assert.AreEqual(expected, model.Model.LastTestInconclusiveCount);
        //    }
        //}

        //[Test]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Succeeded, TestOutcome.Succeeded })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Inconclusive, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Inconclusive, TestOutcome.Inconclusive, TestOutcome.Inconclusive })]
        //[TestCase(new object[] { TestOutcome.Failed, TestOutcome.Failed, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Ignored })]
        //[TestCase(new object[] { TestOutcome.Ignored, TestOutcome.Ignored, TestOutcome.Ignored })]
        //[Category("Unit Testing")]
        //public void LastTestFailedCount_CountIsCorrect(params TestOutcome[] tests)
        //{
        //    var underTest = tests.Select(test => DummyOutcomes[test]).ToList();

        //    using (var model = new MockedTestExplorerModel(underTest))
        //    {
        //        model.Engine.ParserState.OnParseRequested(model);
        //        model.Model.ExecuteTests(model.Model.Tests);
        //        Thread.SpinWait(25);

        //        var expected = tests.Count(outcome => outcome == TestOutcome.Failed);
        //        Assert.AreEqual(expected, model.Model.LastTestFailedCount);
        //    }
        //}

        [Test]
        [Category("Unit Testing")]
        public void CancelTestRun_RequestsCancellation()
        {
            var engine = new Mock <ITestEngine>();
            engine.Setup(m => m.RequestCancellation()).Verifiable("RequestCancellation was not called.");

            using (var model = new TestExplorerModel(engine.Object))
            {
                model.CancelTestRun();
                engine.Verify();
            }
        }

        private static readonly Dictionary<string, Color> ColorLookup = new Dictionary<string, Color>
        {
            { "DimGray", Colors.DimGray },
            { "LimeGreen", Colors.LimeGreen },
            { "Gold", Colors.Gold },
            { "Orange", Colors.Orange },
            { "Red", Colors.Red },
        };

        //[Test]
        //[TestCase("DimGray", new TestOutcome[] { })]
        //[TestCase("Red", new [] { TestOutcome.Succeeded, TestOutcome.Failed })]
        //[TestCase("LimeGreen", new [] { TestOutcome.Succeeded, TestOutcome.Succeeded, TestOutcome.Succeeded })]
        //[TestCase("Red", new [] { TestOutcome.Succeeded, TestOutcome.Inconclusive, TestOutcome.Failed })]
        //[TestCase("Gold", new [] { TestOutcome.Inconclusive, TestOutcome.Inconclusive, TestOutcome.Succeeded })]
        //[TestCase("Red", new [] { TestOutcome.Failed, TestOutcome.Failed, TestOutcome.Failed })]
        //[TestCase("Orange", new [] { TestOutcome.Succeeded, TestOutcome.Ignored })]
        //[Category("Unit Testing")]
        //public void ProgressBarColor_CorrectGivenTestResult(params object[] args)
        //{
        //    var expectedColor = (string)args[0];
        //    var tests = (TestOutcome[])args[1];

        //    var underTest = tests.Select(test => DummyOutcomes[test]).ToList();

        //    using (var model = new MockedTestExplorerModel(underTest))
        //    {
        //        model.Engine.ParserState.OnParseRequested(model);
        //        model.Model.ExecuteTests(model.Model.Tests);
        //        Thread.SpinWait(25);

        //        Assert.AreEqual(ColorLookup[expectedColor], model.Model.ProgressBarColor);
        //    }
        //}
    }
}

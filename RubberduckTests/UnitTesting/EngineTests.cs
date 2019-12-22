using Moq;
using NUnit.Framework;
using Rubberduck.Resources.UnitTesting;
using Rubberduck.UnitTesting;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.UnitTesting
{
    // FIXME - These commented tests need to be restored after TestEngine refactor.
    [NonParallelizable]
    [TestFixture, Apartment(ApartmentState.STA)]  
    public class EngineTests
    {
        [Test]
        [TestCase(1)]
        [TestCase(2)]
        [TestCase(3)]
        [Category("Unit Testing")]
        public void TestEngine_ExposesTestMethods_AndRaisesRefresh(int testCount)
        {
            using (var engine = new MockedTestEngine(testCount))
            {
                var started = 0;
                var refreshes = 0;

                engine.TestEngine.TestsRefreshStarted += (sender, args) => started++;
                engine.TestEngine.TestsRefreshed += (sender, args) => refreshes++;
                engine.ParserState.OnParseRequested(engine);

                if (engine.ParserState.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser Error");
                }

                Assert.AreEqual(1, started);
                Assert.AreEqual(1, refreshes);
                Assert.AreEqual(testCount, engine.TestEngine.Tests.Count());
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_RaisesRefreshEvent_EveryParserRun()
        {
            const int parserRuns = 5;

            using (var engine = new MockedTestEngine(MockedTestEngine.GetTestMethod(1)))
            {
                var refreshCount = 0;
                engine.TestEngine.TestsRefreshed += (sender, args) => refreshCount++;

                for (var parse = 1; parse <= parserRuns; parse++)
                {
                    engine.ParserState.OnParseRequested(engine);

                    if (engine.ParserState.Status != ParserState.Ready)
                    {
                        Assert.Inconclusive("Parser Error");
                    }

                    Assert.AreEqual(parse, refreshCount);
                    Assert.AreEqual(1, engine.TestEngine.Tests.Count());
                }
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_Raises_TestRunStarted()
        {
            SetupAndTestStatusEvent((engine, events) => engine.TestRunStarted += (source, args) => events.Add(args));
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_Raises_TestStarted()
        {
            SetupAndTestStatusEvent((engine, events) => engine.TestStarted += (source, args) => events.Add(args));
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_Raises_TestCompleted()
        {
            SetupAndTestStatusEvent((engine, events) => engine.TestCompleted += (source, args) => events.Add(args));
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_Raises_TestRunCompleted()
        {
            SetupAndTestStatusEvent((engine, events) => engine.TestRunCompleted += (source, args) => events.Add(args));
        }

        private void SetupAndTestStatusEvent(Action<ITestEngine, List<EventArgs>> configuration)
        {
            using (var engine = new MockedTestEngine(MockedTestEngine.GetTestMethod(1)))
            {
                var completionEvents = new List<EventArgs>();

                configuration.Invoke(engine.TestEngine, completionEvents);
                engine.ParserState.OnParseRequested(engine);

                if (engine.ParserState.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser Error");
                }

                engine.TestEngine.Run(engine.TestEngine.Tests);

                Mock.Verify(engine.Dispatcher, engine.VbeInteraction, engine.WrapperProvider);
                Assert.AreEqual(1, completionEvents.Count);
            }
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_RunsAssert_ReturnsSucceeded()
        {
            var expected = new TestResult(TestOutcome.Succeeded);
            SetupAndTestAssertAndReturn(AssertHandler.OnAssertSucceeded, expected);
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_RunsAssert_ReturnsInconclusive()
        {
            var expected = new TestResult(TestOutcome.Inconclusive, "Test Message");
            SetupAndTestAssertAndReturn(() => AssertHandler.OnAssertInconclusive("Test Message"), expected);
        }

        [Test]
        [Category("Unit Testing")]
        public void TestEngine_RunsAssert_ReturnsFailed()
        {
            var expected = new TestResult(TestOutcome.Failed, string.Format(AssertMessages.Assert_FailedMessageFormat, "TestMethod1", "Test Message"));
            // ReSharper disable once ExplicitCallerInfoArgument - there is no "caller".
            SetupAndTestAssertAndReturn(() => AssertHandler.OnAssertFailed("Test Message", "TestMethod1"), expected);
        }

        private void SetupAndTestAssertAndReturn(Action action, TestResult expected)
        {
            using (var engine = new MockedTestEngine(MockedTestEngine.GetTestMethod(1)))
            {
                var completionEvents = new List<TestCompletedEventArgs>();

                engine.SetupAssertCompleted(action);
                engine.TestEngine.TestCompleted += (source, args) => completionEvents.Add(args);
                engine.ParserState.OnParseRequested(engine);

                if (engine.ParserState.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser Error");
                }

                engine.TestEngine.Run(engine.TestEngine.Tests);
                Thread.SpinWait(25);

                Mock.Verify(engine.Dispatcher, engine.VbeInteraction, engine.WrapperProvider, engine.TypeLib);
                Assert.AreEqual(1, completionEvents.Count);
                Assert.AreEqual(expected, completionEvents.First().Result);
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
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Succeeded, TestOutcome.Succeeded })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Inconclusive, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Inconclusive, TestOutcome.Inconclusive, TestOutcome.Inconclusive })]
        //[TestCase(new object[] { TestOutcome.Failed, TestOutcome.Failed, TestOutcome.Failed })]
        //[TestCase(new object[] { TestOutcome.Succeeded, TestOutcome.Ignored })]
        //[TestCase(new object[] { TestOutcome.Ignored, TestOutcome.Ignored, TestOutcome.Ignored })]
        //[Category("Unit Testing")]
        //public void TestEngine_LastTestRun_UpdatesAfterRun(params TestOutcome[] tests)
        //{
        //    var underTest = tests.Select(test => DummyOutcomes[test]).ToList();

        //    using (var engine = new MockedTestEngine(underTest))
        //    {
        //        engine.TestEngine.Run(engine.TestEngine.Tests);
        //        Thread.SpinWait(25);

        //        Assert.AreEqual(underTest.Count, engine.TestEngine.LastRunTests.Count);
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
        //public void TestEngine_RunByOutcome_RunsAppropriateTests(params TestOutcome[] tests)
        //{
        //    var underTest = tests.Select(test => DummyOutcomes[test]).ToList();

        //    using (var engine = new MockedTestEngine(underTest))
        //    {
        //        engine.TestEngine.Run(engine.TestEngine.Tests);

        //        var completionEvents = new List<EventArgs>();
        //        engine.TestEngine.TestCompleted += (source, args) => completionEvents.Add(args);

        //        var outcomes = Enum.GetValues(typeof(TestOutcome)).Cast<TestOutcome>().Where(outcome => outcome != TestOutcome.Unknown);

        //        foreach (var outcome in outcomes)
        //        {
        //            completionEvents.Clear();
        //            engine.TestEngine.RunByOutcome(outcome);

        //            Thread.SpinWait(25);

        //            var expected = tests.Count(result => result == outcome);
        //            Assert.AreEqual(expected, completionEvents.Count);

        //            if (expected == 0)
        //            {
        //                continue;
        //            }

        //            var actual = new List<TestMethod>();
        //            for (var index = 0; index < underTest.Count; index++)
        //            {
        //                if (tests[index] == outcome)
        //                {
        //                    actual.Add(engine.TestEngine.Tests.ElementAt(index));
        //                }
        //            }

        //            CollectionAssert.AreEqual(actual, engine.TestEngine.LastRunTests);
        //        }
        //    }
        //}

    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Rubberduck.UI.UnitTesting;
using Rubberduck.UnitTesting;

namespace RubberduckTests.UnitTesting
{
    internal class MockedTestExplorerModel : IDisposable
    {
        public MockedTestExplorerModel(MockedTestEngine engine)
        {
            Engine = engine;
            Model = new TestExplorerModel(Engine.TestEngine);
        }

        public MockedTestExplorerModel(List<(TestOutcome Outcome, string Output, long duration)> results)
        {
            var code = string.Join(Environment.NewLine,
                Enumerable.Range(1, results.Count).Select(num =>
                    MockedTestEngine.GetTestMethod(num, results[num - 1].Outcome == TestOutcome.Ignored)));

            Engine = new MockedTestEngine(code);
            Model = new TestExplorerModel(Engine.TestEngine);
            Engine.ParserState.OnParseRequested(this);

            var testMethodCount = results.Count;
            var testMethods = Engine.TestEngine.Tests.ToList();

            if (testMethods.Count != testMethodCount)
            {
                Assert.Inconclusive("Test setup failure.");
            }

            for (var test = 0; test < results.Count; test++)
            {
                var (outcome, output, duration) = results[test];
                Engine.SetupAssertCompleted(testMethods[test], new TestResult(outcome, output, duration));
            }
        }

        public TestExplorerModel Model { get; set; }

        public MockedTestEngine Engine { get; set; }

        public void Dispose()
        {
            Engine?.Dispose();
            Model?.Dispose();
        }
    }
}

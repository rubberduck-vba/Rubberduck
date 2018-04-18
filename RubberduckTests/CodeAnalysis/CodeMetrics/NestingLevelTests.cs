using NUnit.Framework;
using Rubberduck.CodeAnalysis.CodeMetrics;
using RubberduckTests.Mocks;
using System;
using System.Linq;
using System.Text;

namespace RubberduckTests.CodeAnalysis.CodeMetrics
{
    [TestFixture]
    public class NestingLevelTests
    {
        private CodeMetricsAnalyst analyst;

        [SetUp]
        public void Setup()
        {
            analyst = new CodeMetricsAnalyst(new CodeMetric[] { new NestingLevelMetric() });
        }

        [Test]
        [Category("Code Metrics")]
        public void EmtpyModule_HasNoNestingLevelMetrics()
        {
            var code = @"";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metrics = analyst.GetMetrics(state).ToList();
                Assert.AreEqual(0, metrics.Count);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void SimpleSub_HasNestingLevel_One()
        {
            var code = @"
Option Explicit

Public Sub SimpleSub()
    'preceding comment just to check
    Debug.Print ""this is a test""
End Sub
";

            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("1", metricResult.Value);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void WeirdSub_HasNestingLevel_One()
        {
            var code = @"
Option Explicit

Public Sub WeirdSub()
    ' some comments
    Debug.Print ""An expression, that "" & _
            ""extends across multiple lines, with "" _
                & ""Line continuations that do weird stuff "" & _
         ""but shouldn't account for nesting""
    Debug.Print ""Just to confuse you""
End Sub
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("1", metricResult.Value);
            }
        }

    }
}

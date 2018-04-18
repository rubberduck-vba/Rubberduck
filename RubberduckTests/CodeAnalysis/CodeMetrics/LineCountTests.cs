using NUnit.Framework;
using Rubberduck.CodeAnalysis.CodeMetrics;
using RubberduckTests.Mocks;
using System;
using System.Linq;
using System.Text;

namespace RubberduckTests.CodeAnalysis.CodeMetrics
{
    class LineCountTests
    {
        private CodeMetricsAnalyst analyst;

        [SetUp]
        public void Setup()
        {
            analyst = new CodeMetricsAnalyst(new[]
            {
                new LineCountModuleMetric()
            });
        }

        [Test]
        [Category("Code Metrics")]
        public void EmptyModule_HasMetricsZeroed()
        {
            var code = @"";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("0", metricResult.Value);
            }
        }


        [Test]
        [Category("Code Metrics")]
        public void ModuleHas_AsManyLines_AsPhysicalLines()
        {
            foreach (var lineCount in new int[] { 0, 10, 15, 200, 1020 })
            {
                var builder = new StringBuilder();
                for (int i = 0; i < lineCount; i++)
                {
                    builder.Append(Environment.NewLine);
                }
                var code = builder.ToString();

                using (var state = MockParser.ParseString(code, out var _))
                {
                    var metricResult = analyst.GetMetrics(state).First();
                    Assert.AreEqual(lineCount.ToString(), metricResult.Value);
                }
            }
        }
    }
}

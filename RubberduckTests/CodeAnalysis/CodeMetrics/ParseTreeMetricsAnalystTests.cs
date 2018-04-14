using NUnit.Framework;
using Rubberduck.CodeAnalysis.CodeMetrics;
using RubberduckTests.Mocks;
using System;
using System.Linq;
using System.Text;

namespace RubberduckTests.CodeAnalysis.CodeMetrics
{
    [TestFixture]
    public class CodeMetricsAnalystTests
    {
        private CodeMetricsAnalyst cut;

        [SetUp]
        public void Setup()
        {
            cut = new CodeMetricsAnalyst(new CodeMetric[] { });
        }

        [Test, Ignore("under rewrite")]
        [Category("Code Metrics")]
        public void EmptyModule_HasMetricsZeroed()
        {
            var code = @"";
            var state = MockParser.ParseString(code, out var qmn);
            var metrics = cut.GetMetrics(state).First();
            //Assert.AreEqual(new CodeMetricsResult(), metrics.Result);
        }

        [Test, Ignore("under rewrite")]
        [Category("Code Metrics")]
        public void ModuleHas_AsManyLines_AsPhysicalLines()
        {
            foreach (var lineCount in new int[]{ 0, 10, 15, 200, 1020 })
            {
                var builder = new StringBuilder();
                for (int i = 0; i < lineCount; i++)
                {
                    builder.Append(Environment.NewLine);
                }
                var code = builder.ToString();

                var state = MockParser.ParseString(code, out var _);
                var metric = cut.GetMetrics(state).First();
                //Assert.AreEqual(lineCount, metric.Result.Lines);
            }
        }


        [Test, Ignore("under rewrite")]
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

            var state = MockParser.ParseString(code, out var _);
            var metrics = cut.GetMetrics(state).First();
            //Assert.AreEqual(1, metrics.Result.MaximumNesting);
        }

        [Test, Ignore("under rewrite")]
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
                var metrics = cut.GetMetrics(state).First();
                //Assert.AreEqual(1, metrics.Result.MaximumNesting);
            }
        }

    }
}

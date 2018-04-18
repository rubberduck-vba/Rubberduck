using NUnit.Framework;
using Rubberduck.CodeAnalysis.CodeMetrics;
using RubberduckTests.Mocks;
using System.Linq;
using System.Text;

namespace RubberduckTests.CodeAnalysis.CodeMetrics
{
    [TestFixture]
    public class CyclomaticComplexityTests
    {

        private CodeMetricsAnalyst analyst;

        [SetUp]
        public void Setup()
        {
            analyst = new CodeMetricsAnalyst(new[]
            {
                new CyclomaticComplexityMemberMetric()
            });
        }

        [Test]
        [Category("Code Metrics")]
        public void EmptyModule_HasNoMemberMetricResults()
        {
            var code = @"";
            using (var state = MockParser.ParseString(code, out var qmn))
            {
                var metrics = analyst.GetMetrics(state).ToList();
                Assert.AreEqual(0, metrics.Count);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void EmptySub_HasCyclomaticComplexity_One()
        {
            var code = @"
Sub NoCode()
End Sub
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metrics = analyst.GetMetrics(state).ToList();
                Assert.AreEqual(1, metrics.Count);
                Assert.AreEqual("1", metrics[0].Value);
                Assert.AreEqual(state.DeclarationFinder.UserDeclarations(Rubberduck.Parsing.Symbols.DeclarationType.Procedure).First()
                    , metrics[0].Declaration);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void EmptyFunction_HasCyclomaticComplexit_One()
        {
            var code = @"
Function NoCode()
End Function
";

            using (var state = MockParser.ParseString(code, out var _))
            {
                var metric = analyst.GetMetrics(state).First();
                Assert.AreEqual("1", metric.Value);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void SingleIfStatement_HasCyclomaticComplexity_2()
        {
            var code = @"
Sub IfStatement()
    If True Then
    End If
End Sub
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("2", metricResult.Value);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void SingleIfElseStatement_HasCyclomaticComplexity_2()
        {
            var code = @"
Sub IfElseStatement()
    If True Then
    Else
    End If
End Sub
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("2", metricResult.Value);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void IfElseIfStatement_HasCyclomaticComplexity_3()
        {
            var code = @"
Sub IfElseifStatement()
    If True Then
    ElseIf False Then
    End If
End Sub
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("3", metricResult.Value);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void IfElseIfElseStatement_HasCyclomaticComplexity_3()
        {
            var code = @"
Sub IfElseifStatement()
    If True Then
    ElseIf False Then
    Else
    End If
End Sub
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("3", metricResult.Value);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void NestedIfStatement_HasCyclomaticComplexity_3()
        {
            var code = @"
Sub IfElseifStatement()
    If True Then
        If False Then
        End If
    End If
End Sub
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("3", metricResult.Value);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void ForeachLoop_HasCyclomaticComplexity_2()
        {
            var code = @"
Sub ForeachLoop(ByRef iterable As Object)
    Dim stuff As Variant
    For Each stuff In iterable 
    Next stuff
End Sub
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("2", metricResult.Value);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void ForToNextLoop_HasCyclomaticComplexity_2()
        {
            var code = @"
Sub ForToNextLoop(ByVal ubound As Long)
    Dim i As Long
    For i = 0 To ubound Step 1
        ' nothing
    Next i
End Sub
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("2", metricResult.Value);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void CaseOnlyElse_HasCyclomaticComplexity_1()
        {
            var code = @"
Sub CaseOnlyElse(ByVal number As Long) 
    Select Case number
        Case Else
    End Select
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
        public void CaseBlock_HasCyclomaticComplexity_CorrespondingToCaseLabels()
        {
            foreach (var blockCount in new int[] { 1, 2, 5, 25, 40 })
            {
                var caseBlockBuilder = new StringBuilder();
                for (int i = 0; i < blockCount; i++)
                {
                    caseBlockBuilder.Append($"\r\n        Case number < {i}\r\n\r\n");
                }
                var code = @"
Sub CaseBlockWithCounts(ByVal number As Long)
    Select Case number
" + caseBlockBuilder.ToString() + @"
        Case Else
    End Select
End Sub
";
                using (var state = MockParser.ParseString(code, out var _))
                {
                    var metricResult = analyst.GetMetrics(state).First();
                    Assert.AreEqual((blockCount + 1).ToString(), metricResult.Value);
                }
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void PropertyGet_HasCyclomaticComplexity_One()
        {
            var code = @"
Public Property Get Complexity() As Long
    Complexity = 1
End Property
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("1", metricResult.Value);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void PropertyLet_HasCyclomaticComplexity_One()
        {
            var code = @"
Option Explicit

Private mComplexity As Long

Public Property Let Complexity(ByVal complexity As Long)
    mComplexity = complexity
End Property
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("1", metricResult.Value);
            }
        }

        [Test]
        [Category("Code Metrics")]
        public void PropertySet_HasCyclomaticComplexity_One()
        {
            var code = @"
Option Explicit

Private mComplexity As Object

Public Property Set Complexity(ByRef complexity As Object)
    mComplexity = complexity
End Property
";
            using (var state = MockParser.ParseString(code, out var _))
            {
                var metricResult = analyst.GetMetrics(state).First();
                Assert.AreEqual("1", metricResult.Value);
            }
        }


    }
}

using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UnassignedVariableUsageInspectionTests
    {
        private IEnumerable<IInspectionResult> GetInspectionResults(string code)
        {
            var vbe = MockVbeBuilder.BuildFromSingleModule(code, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnassignedVariableUsageInspection(state);
                return inspection.GetInspectionResults(CancellationToken.None);
            }
        }

        [Test]
        [Category("Inspections")]
        public void IgnoresExplicitArrays()
        {
            const string code = @"
Sub Foo()
    Dim bar() As String
    bar(1) = ""value""
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void IgnoresArrayReDim()
        {
            const string code = @"
Sub Foo()
    Dim bar As Variant
    ReDim bar(1 To 10)
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void IgnoresArraySubscripts()
        {
            const string code = @"
Sub Foo()
    Dim bar As Variant
    ReDim bar(1 To 10)
    bar(1) = 42
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariableUsage_ReturnsResult()
        {
            const string code = @"
Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean
    bb = b
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariableUsage_DoesNotReturnResult()
        {
            const string code = @"
Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean
    b = True
    bb = b
End Sub
";

            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariableUsage_Ignored_DoesNotReturnResult()
        {
            const string code = @"
Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean

'@Ignore UnassignedVariableUsage
    bb = b
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariableUsage_Ignored_DoesNotReturnResultMultipleIgnores()
        {
            const string code = @"
Sub Foo()    
    Dim b As Boolean
    Dim bb As Boolean

'@Ignore UnassignedVariableUsage, VariableNotAssigned
    bb = b
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void UnassignedVariableUsage_NoResultIfNoReferences()
        {
            const string code = @"
Sub DoSomething()
    Dim foo
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Ignore("Test concurrency issue. Only passes if run individually.")]
        [Category("Inspections")]
        public void UnassignedVariableUsage_NoResultForLenFunction()
        {
            const string code = @"
Sub DoSomething()
    Dim foo As LongPtr
    Debug.Print Len(foo)
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Ignore("Test concurrency issue. Only passes if run individually.")]
        [Category("Inspections")]
        public void UnassignedVariableUsage_NoResultForLenBFunction()
        {
            const string code = @"
Sub DoSomething()
    Dim foo As LongPtr
    Debug.Print LenB(foo)
End Sub
";
            var results = GetInspectionResults(code);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "UnassignedVariableUsageInspection";
            var inspection = new UnassignedVariableUsageInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}

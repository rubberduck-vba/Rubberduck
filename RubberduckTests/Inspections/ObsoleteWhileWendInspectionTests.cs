using NUnit.Framework;
using Rubberduck.Parsing.Inspections.Abstract;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.CodeAnalysis.Inspections.Concrete;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteWhileWendInspectionTests
    {
        private IEnumerable<IInspectionResult> Inspect(string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new ObsoleteWhileWendStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                return inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteWhileWendLoop_NoWhileWendLoop_NoResult()
        {
            const string inputCode = @"
Sub Foo()
    Do While True
    Loop
End Sub
";
            var results = Inspect(inputCode);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteWhileWendLoop_IgnoredWhileWendLoop_NoResult()
        {
            const string inputCode = @"
Sub Foo()
    '@Ignore ObsoleteWhileWendStatement
    While True
    Wend
End Sub
";
            var results = Inspect(inputCode);
            Assert.AreEqual(0, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteWhileWendLoop_EmptyBody_ReturnsResult()
        {
            const string inputCode = @"
Sub Foo()
    While True
    Wend
End Sub
";
            var results = Inspect(inputCode);
            Assert.AreEqual(1, results.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteWhileWendLoop_NonEmptyBody_ReturnsResult()
        {
            const string inputCode = @"
Sub Foo()
    Dim bar As Long
    While bar < 12
        bar = bar + 1
    Wend
End Sub
";
            var results = Inspect(inputCode);
            Assert.AreEqual(1, results.Count());
        }
    }
}
using NUnit.Framework;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteWhileWendInspectionTests : InspectionTestsBase
    {
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
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
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
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
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
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
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
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ObsoleteWhileWendStatementInspection(state);
        }
    }
}
using System.Linq;
using NUnit.Framework;
using RubberduckTests.Mocks;
using System.Threading;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ObsoleteCallStatementInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo()
    Call Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
    Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_DoesNotReturnResult_InstructionSeparator()
        {
            const string inputCode =
                @"Sub Foo()
    Call Foo: Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsResult_ColonInComment()
        {
            const string inputCode =
                @"Sub Foo()
    Call Foo ' I''ve got a colon: see?
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsResult_ColonInStringLiteral()
        {
            const string inputCode =
                @"Sub Foo(ByVal str As String)
    Call Foo("":"")
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsMultipleResults()
        {
            const string inputCode =
                @"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Call Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_ReturnsResults_SomeObsoleteCallStatements()
        {
            const string inputCode =
                @"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ObsoleteCallStatement_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
    '@Ignore ObsoleteCallStatement
    Call Foo
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }


        [Test]
        [Category("Inspections")]
        public void InspectionType()
        {
            var inspection = new ObsoleteCallStatementInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteCallStatementInspection";
            var inspection = new ObsoleteCallStatementInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}

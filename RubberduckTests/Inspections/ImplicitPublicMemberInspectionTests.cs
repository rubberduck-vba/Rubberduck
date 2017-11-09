using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ImplicitPublicMemberInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitPublicMember_ReturnsResult_Sub()
        {
            const string inputCode = @"
Sub Foo()
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitPublicMemberInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitPublicMember_ReturnsResult_Function()
        {
            const string inputCode =
                @"Function Foo() As Boolean
    Foo = True
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitPublicMemberInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitPublicMember_ReturnsResult_MultipleSubs()
        {
            const string inputCode =
                @"Sub Foo()
End Sub

Sub Goo()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitPublicMemberInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitPublicMember_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private Sub Foo()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitPublicMemberInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitPublicMember_Ignored_DoesNotReturnResult_Sub()
        {
            const string inputCode = @"
'@Ignore ImplicitPublicMember
Sub Foo()
End Sub
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitPublicMemberInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitPublicMember_ReturnsResult_SomeImplicitlyPublicSubs()
        {
            const string inputCode =
                @"Private Sub Foo()
End Sub

Sub Goo()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitPublicMemberInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ImplicitPublicMemberInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitPublicMemberInspection";
            var inspection = new ImplicitPublicMemberInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}

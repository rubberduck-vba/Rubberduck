using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class WriteOnlyPropertyQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void WriteOnlyProperty_AddPropertyGetQuickFixWorks_ImplicitTypesAndAccessibility()
        {
            const string inputCode =
                @"Property Let Foo(value)
End Property";

            const string expectedCode =
                @"Public Property Get Foo() As Variant
End Property

Property Let Foo(value)
End Property";


            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out var component);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new WriteOnlyPropertyInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new WriteOnlyPropertyQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void WriteOnlyProperty_AddPropertyGetQuickFixWorks_ExlicitTypesAndAccessibility()
        {
            const string inputCode =
                @"Public Property Let Foo(ByVal value As Integer)
End Property";

            const string expectedCode =
                @"Public Property Get Foo() As Integer
End Property

Public Property Let Foo(ByVal value As Integer)
End Property";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out var component);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new WriteOnlyPropertyInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new WriteOnlyPropertyQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void WriteOnlyProperty_AddPropertyGetQuickFixWorks_MultipleParams()
        {
            const string inputCode =
                @"Public Property Let Foo(value1, ByVal value2 As Integer, ByRef value3 As Long, value4 As Date, ByVal value5, value6 As String)
End Property";

            const string expectedCode =
                @"Public Property Get Foo(ByRef value1 As Variant, ByVal value2 As Integer, ByRef value3 As Long, ByRef value4 As Date, ByVal value5 As Variant) As String
End Property

Public Property Let Foo(value1, ByVal value2 As Integer, ByRef value3 As Long, value4 As Date, ByVal value5, value6 As String)
End Property";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out var component);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new WriteOnlyPropertyInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new WriteOnlyPropertyQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}

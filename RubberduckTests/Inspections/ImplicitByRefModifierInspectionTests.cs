using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplicitByRefModifierInspectionTests : InspectionTestsBase
    {
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_ReturnsResult_MultipleParameters()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Integer, arg2 As Date)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_DoesNotReturnResult_ByRef()
        {
            const string inputCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_DoesNotReturnResult_ByVal()
        {
            const string inputCode =
                @"Sub Foo(ByVal arg1 As Integer)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_ReturnsResult_SomePassedByRefImplicitly()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Integer, ByRef arg2 As Date)
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_DoesNotReturnResult_ParamArray()
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_ReturnsResult_InterfaceImplementation()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";


            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_ReturnsResult_MultipleInterfaceImplementations()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            const string inputCode3 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ImplicitByRefModifier
Sub Foo(arg1 As Integer)
End Sub";
        }

        [Test]
        [Category("QuickFixes")]
        public void InspectionName()
        {
            var inspection = new ImplicitByRefModifierInspection(null);

        }
    }
}

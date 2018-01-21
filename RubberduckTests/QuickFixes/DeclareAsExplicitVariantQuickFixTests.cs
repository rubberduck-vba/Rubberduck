using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class DeclareAsExplicitVariantQuickFixTests
    {

        [Test]
        [Category("QuickFixes")]
        public void VariableTypeNotDeclared_QuickFixWorks_Parameter()
        {
            const string inputCode =
                @"Sub Foo(arg1)
End Sub";

            const string expectedCode =
                @"Sub Foo(arg1 As Variant)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                new DeclareAsExplicitVariantQuickFix(state).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void VariableTypeNotDeclared_QuickFixWorks_SubNameContainsParameterName()
        {
            const string inputCode =
                @"Sub Foo(Foo)
End Sub";

            const string expectedCode =
                @"Sub Foo(Foo As Variant)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                new DeclareAsExplicitVariantQuickFix(state).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void VariableTypeNotDeclared_QuickFixWorks_Variable()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim var1 As Variant
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                new DeclareAsExplicitVariantQuickFix(state).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void VariableTypeNotDeclared_QuickFixWorks_ParameterWithoutDefaultValue()
        {
            const string inputCode =
                @"Sub Foo(ByVal Fizz)
End Sub";

            const string expectedCode =
                @"Sub Foo(ByVal Fizz As Variant)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                new DeclareAsExplicitVariantQuickFix(state).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void VariableTypeNotDeclared_QuickFixWorks_ParameterWithDefaultValue()
        {
            const string inputCode =
                @"Sub Foo(ByVal Fizz = False)
End Sub";

            const string expectedCode =
                @"Sub Foo(ByVal Fizz As Variant = False)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                new DeclareAsExplicitVariantQuickFix(state).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}

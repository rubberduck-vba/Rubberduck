using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class DeclareAsExplicitVariantQuickFixTests
    {

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void VariableTypeNotDeclared_QuickFixWorks_Parameter()
        {
            const string inputCode =
@"Sub Foo(arg1)
End Sub";

            const string expectedCode =
@"Sub Foo(arg1 As Variant)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableTypeNotDeclaredInspection(state);
            new DeclareAsExplicitVariantQuickFix(state).Fix(inspection.GetInspectionResults().First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void VariableTypeNotDeclared_QuickFixWorks_SubNameContainsParameterName()
        {
            const string inputCode =
@"Sub Foo(Foo)
End Sub";

            const string expectedCode =
@"Sub Foo(Foo As Variant)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableTypeNotDeclaredInspection(state);
            new DeclareAsExplicitVariantQuickFix(state).Fix(inspection.GetInspectionResults().First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableTypeNotDeclaredInspection(state);
            new DeclareAsExplicitVariantQuickFix(state).Fix(inspection.GetInspectionResults().First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void VariableTypeNotDeclared_QuickFixWorks_ParameterWithoutDefaultValue()
        {
            const string inputCode =
@"Sub Foo(ByVal Fizz)
End Sub";

            const string expectedCode =
@"Sub Foo(ByVal Fizz As Variant)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableTypeNotDeclaredInspection(state);
            new DeclareAsExplicitVariantQuickFix(state).Fix(inspection.GetInspectionResults().First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void VariableTypeNotDeclared_QuickFixWorks_ParameterWithDefaultValue()
        {
            const string inputCode =
@"Sub Foo(ByVal Fizz = False)
End Sub";

            const string expectedCode =
@"Sub Foo(ByVal Fizz As Variant = False)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new VariableTypeNotDeclaredInspection(state);
            new DeclareAsExplicitVariantQuickFix(state).Fix(inspection.GetInspectionResults().First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

    }
}

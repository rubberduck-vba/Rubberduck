using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class DeclareAsExplicitVariantQuickFixTests :  QuickFixTestBase
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
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

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new VariableTypeNotDeclaredInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new DeclareAsExplicitVariantQuickFix();
        }
    }
}

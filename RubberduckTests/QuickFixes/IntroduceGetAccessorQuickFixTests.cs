using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class IntroduceGetAccessorQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void IntroduceGetAccessor_AddPropertyGetQuickFixWorks_ImplicitTypesAndAccessibility()
        {
            const string inputCode =
                @"Property Let Foo(value)
End Property";

            const string expectedCode =
                @"Public Property Get Foo() As Variant
End Property

Property Let Foo(value)
End Property";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new WriteOnlyPropertyInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntroduceGetAccessor_AddPropertyGetQuickFixWorks_ExlicitTypesAndAccessibility()
        {
            const string inputCode =
                @"Public Property Let Foo(ByVal value As Integer)
End Property";

            const string expectedCode =
                @"Public Property Get Foo() As Integer
End Property

Public Property Let Foo(ByVal value As Integer)
End Property";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new WriteOnlyPropertyInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IntroduceGetAccessor_AddPropertyGetQuickFixWorks_MultipleParams()
        {
            const string inputCode =
                @"Public Property Let Foo(value1, ByVal value2 As Integer, ByRef value3 As Long, value4 As Date, ByVal value5, value6 As String)
End Property";

            const string expectedCode =
                @"Public Property Get Foo(ByRef value1 As Variant, ByVal value2 As Integer, ByRef value3 As Long, ByRef value4 As Date, ByVal value5 As Variant) As String
End Property

Public Property Let Foo(value1, ByVal value2 As Integer, ByRef value3 As Long, value4 As Date, ByVal value5, value6 As String)
End Property";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new WriteOnlyPropertyInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new IntroduceGetAccessorQuickFix();
        }

        protected override IVBE TestVbe(string code, out IVBComponent component)
        {
            return MockVbeBuilder.BuildFromSingleModule(code, ComponentType.ClassModule, out component).Object;
        }
    }
}

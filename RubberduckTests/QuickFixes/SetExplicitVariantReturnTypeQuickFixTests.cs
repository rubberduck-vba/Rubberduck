using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class SetExplicitVariantReturnTypeQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void ImplicitVariantReturnType_QuickFixWorks_Function()
        {
            const string inputCode =
                @"Function Foo()
End Function";

            const string expectedCode =
                @"Function Foo() As Variant
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ImplicitVariantReturnTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitVariantReturnType_QuickFixWorks_PropertyGet()
        {
            const string inputCode =
                @"Property Get Foo()
End Property";

            const string expectedCode =
                @"Property Get Foo() As Variant
End Property";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ImplicitVariantReturnTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitVariantReturnType_QuickFixWorks_LibraryFunction()
        {
            const string inputCode =
                @"Declare PtrSafe Function CreateProcess Lib ""kernel32"" _
                                   Alias ""CreateProcessA""(ByVal lpApplicationName As String, _
                                                           ByVal lpCommandLine As String, _
                                                           lpProcessAttributes As SECURITY_ATTRIBUTES, _
                                                           lpThreadAttributes As SECURITY_ATTRIBUTES, _
                                                           ByVal bInheritHandles As Long, _
                                                           ByVal dwCreationFlags As Long, _
                                                           lpEnvironment As Any, _
                                                           ByVal lpCurrentDirectory As String, _
                                                           lpStartupInfo As STARTUPINFO, _
                                                           lpProcessInformation As PROCESS_INFORMATION)";

            const string expectedCode =
                @"Declare PtrSafe Function CreateProcess Lib ""kernel32"" _
                                   Alias ""CreateProcessA""(ByVal lpApplicationName As String, _
                                                           ByVal lpCommandLine As String, _
                                                           lpProcessAttributes As SECURITY_ATTRIBUTES, _
                                                           lpThreadAttributes As SECURITY_ATTRIBUTES, _
                                                           ByVal bInheritHandles As Long, _
                                                           ByVal dwCreationFlags As Long, _
                                                           lpEnvironment As Any, _
                                                           ByVal lpCurrentDirectory As String, _
                                                           lpStartupInfo As STARTUPINFO, _
                                                           lpProcessInformation As PROCESS_INFORMATION) As Variant";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ImplicitVariantReturnTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitVariantReturnType_QuickFixWorks_Function_HasComment()
        {
            const string inputCode =
                @"Function Foo()    ' comment
End Function";

            const string expectedCode =
                @"Function Foo() As Variant    ' comment
End Function";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new ImplicitVariantReturnTypeInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new SetExplicitVariantReturnTypeQuickFix();
        }
    }
}

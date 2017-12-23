using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class SetExplicitVariantReturnTypeQuickFixTests
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitVariantReturnTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new SetExplicitVariantReturnTypeQuickFix(state).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitVariantReturnTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new SetExplicitVariantReturnTypeQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitVariantReturnTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new SetExplicitVariantReturnTypeQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitVariantReturnTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new SetExplicitVariantReturnTypeQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }
    }
}

using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestClass]
    public class SetExplicitVariantReturnTypeQuickFixTests
    {
        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ImplicitVariantReturnType_QuickFixWorks_Function()
        {
            const string inputCode =
@"Function Foo()
End Function";

            const string expectedCode =
@"Function Foo() As Variant
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitVariantReturnTypeInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new SetExplicitVariantReturnTypeQuickFix(state).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ImplicitVariantReturnType_QuickFixWorks_PropertyGet()
        {
            const string inputCode =
@"Property Get Foo()
End Property";

            const string expectedCode =
@"Property Get Foo() As Variant
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitVariantReturnTypeInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new SetExplicitVariantReturnTypeQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
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

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitVariantReturnTypeInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new SetExplicitVariantReturnTypeQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("QuickFixes")]
        public void ImplicitVariantReturnType_QuickFixWorks_Function_HasComment()
        {
            const string inputCode =
@"Function Foo()    ' comment
End Function";

            const string expectedCode =
@"Function Foo() As Variant    ' comment
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitVariantReturnTypeInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new SetExplicitVariantReturnTypeQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }
    }
}

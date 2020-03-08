using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplicitVariantReturnTypeInspectionTests : InspectionTestsBase
    {
        [TestCase("Function Foo()\r\nEnd Function", 1)]
        [TestCase("Function Foo() As Boolean\r\nEnd Function", 0)]
        [TestCase("Property Get Foo()\r\nEnd Property", 1)]
        [TestCase("Function Foo()\r\nEnd Function\r\n\r\nFunction Goo()\r\nEnd Function", 2)]
        [TestCase("Function Foo()\r\nEnd Function\r\n\r\nFunction Goo() As String\r\nEnd Function", 1)]
        [TestCase("Property Get Foo()\r\nEnd Property", 1)]
        [TestCase("'@Ignore ImplicitVariantReturnType\r\n\r\nFunction Foo()\r\nEnd Function", 0)]
        [Category("Inspections")]
        public void ImplicitVariantReturnType_VariousScenarios(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitVariantReturnType_ReturnsResult_LibraryFunction()
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
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitVariantReturnType_DoesNotReturnResult_LibraryFunction()
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
                                                           lpProcessInformation As PROCESS_INFORMATION) As Long";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ImplicitVariantReturnTypeInspection(null);

            Assert.AreEqual(nameof(ImplicitVariantReturnTypeInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplicitVariantReturnTypeInspection(state);
        }
    }
}

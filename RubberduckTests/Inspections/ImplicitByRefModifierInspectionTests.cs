using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplicitByRefModifierInspectionTests : InspectionTestsBase
    {
        [TestCase("Sub Foo(arg1 As Integer)\r\nEnd Sub", 1)]
        [TestCase("Sub Foo(arg1 As Integer, arg2 As Date)\r\nEnd Sub", 2)]
        [TestCase("Sub Foo(ByRef arg1 As Integer)\r\nEnd Sub", 0)]
        [TestCase("Sub Foo(ByVal arg1 As Integer)\r\nEnd Sub", 0)]
        [TestCase("Sub Foo(arg1 As Integer, ByRef arg2 As Date)\r\nEnd Sub", 1)]
        [TestCase("Sub Foo(ParamArray arg1 As Integer)\r\nEnd Sub", 0)]
        [Category("QuickFixes")]
        [Category(nameof(ImplicitByRefModifierInspection))]
        public void ImplicitByRefModifier_SimpleScenarios(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [TestCase("Property Let Fizz(RHS As Integer)\r\nEnd Property", 0)]
        [TestCase("Property Set Fizz(RHS As Object)\r\nEnd Property", 0)]
        [TestCase("Property Let Fizz(index As Integer, RHS As Integer)\r\nEnd Property", 1)]
        [TestCase("Property Set Fizz(index As Integer, RHS As Object)\r\nEnd Property", 1)]
        [Category("QuickFixes")]
        [Category(nameof(ImplicitByRefModifierInspection))]
        public void ImplicitByRefModifier_PropertyMutatorRHSParameter(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(ImplicitByRefModifierInspection))]
        public void ImplicitByRefModifier_ReturnsResult_InterfaceImplementation()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
           {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(ImplicitByRefModifierInspection))]
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
        [Category(nameof(ImplicitByRefModifierInspection))]
        public void ImplicitByRefModifier_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ImplicitByRefModifier
Sub Foo(arg1 As Integer)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [TestCase(@"Public Declare PtrSafe Sub LibProcedure Lib ""MyLib""(arg As Long)", "LibProcedure 2000")]
        [TestCase(@"Public Declare PtrSafe Function LibProcedure Lib ""MyLib""(arg As Long) As Long", "test = LibProcedure(2000)")]
        [Category("QuickFixes")]
        [Category(nameof(ImplicitByRefModifierInspection))]
        public void ImplicitByRefModifier_IgnoresDeclareStatement(string declareStatement, string libraryCall)
        {
            var inputCode =
$@"
Option Explicit

Private test As Long

{declareStatement}

Public Sub CallTheLib()
    {libraryCall}
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(ImplicitByRefModifierInspection))]
        public void InspectionName()
        {
            var inspection = new ImplicitByRefModifierInspection(null);

            Assert.AreEqual(nameof(ImplicitByRefModifierInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplicitByRefModifierInspection(state);
        }
    }
}

using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class IntegerDataTypeInspectionTests : InspectionTestsBase
    {
        [TestCase("Function Foo() As Integer\r\nEnd Function")]
        [TestCase("Property Get Foo() As Integer\r\nEnd Property")]
        [TestCase("Sub Foo(arg As Integer)\r\nEnd Sub")]
        [TestCase("Sub Foo()\r\nDim v as Integer\r\nEnd Sub")]
        [TestCase("Sub Foo()\r\nConst c As Integer = 0\r\nEnd Sub")]
        [TestCase("Type T\r\ni As Integer\r\nEnd Type")]
        [Category("Inspections")]
        public void IntegerDataType_Various_ReturnsResults(string inputCode)
        {
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void IntegerDataType_ReturnsResult_FunctionInterfaceImplementation()
        {
            const string inputCode1 =
                @"Function Foo() As Integer
End Function";

            const string inputCode2 =
                @"Implements IClass1

Function IClass1_Foo() As Integer
End Function";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void IntegerDataType_ReturnsResult_PropertyGetInterfaceImplementation()
        {
            const string inputCode1 =
                @"Property Get Foo() As Integer
End Property";

            const string inputCode2 =
                @"Implements IClass1

Property Get IClass1_Foo() As Integer
End Property";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void IntegerDataType_ReturnsResult_ParameterInterfaceImplementation()
        {
            const string inputCode1 =
                @"Sub Foo(arg as Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void IntegerDataType_ReturnsResult_MultipleInterfaceImplementations()
        {
            const string inputCode1 =
                @"Sub Foo(arg As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg As Integer)
End Sub";

            const string inputCode3 =
                @"Implements IClass1

Sub IClass1_Foo(arg As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[] 
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [TestCase(@"Declare Function Foo Lib ""lib.dll"" () As Integer")] //ignores library function elements
        [TestCase(@"Declare Function Foo Lib ""lib.dll"" (arg As Integer) As String")] //ignores library function elements
        [TestCase(@"Declare Sub Foo Lib ""lib.dll"" (arg As Integer)")] //ignores library function elements
        [TestCase("'@Ignore IntegerDataType\r\nSub Foo(arg1 As Integer)\r\nEnd Sub")]
        [Category("Inspections")]
        public void IntegerDataType_Various_NoResults(string inputCode)
        {
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new IntegerDataTypeInspection(null);

            Assert.AreEqual(nameof(IntegerDataTypeInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new IntegerDataTypeInspection(state);
        }
    }
}

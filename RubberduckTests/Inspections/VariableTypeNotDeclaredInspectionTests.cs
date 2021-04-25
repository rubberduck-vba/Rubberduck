using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class VariableTypeNotDeclaredInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_Parameter()
        {
            const string inputCode =
                @"Sub Foo(arg1)
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_MultipleParams()
        {
            const string inputCode =
                @"Sub Foo(arg1, arg2)
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_DoesNotReturnResult_Parameter()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Date)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_SomeTypesNotDeclared_Parameters()
        {
            const string inputCode =
                @"Sub Foo(arg1, arg2 As String)
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_SomeTypesNotDeclared_Variables()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1
    Dim var2 As Date
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_Variable()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_ReturnsResult_MultipleVariables()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1
    Dim var2
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_DoesNotReturnResult_Variable()
        {
            const string inputCode =
                @"Sub Foo()
    Dim var1 As Integer
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore VariableTypeNotDeclared
Sub Foo(arg1)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableTypeNotDeclared_Const_DoesNotReturnResult()
        {
            const string inputCode =
                @"Sub Foo()
    Const bar = 42
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Variant")]
        [TestCase("Long")]
        public void VariableTypeNotDeclared_TypedArray_Dim_DoesNotReturnResult(string variableType)
        {
            var inputCode =
                $@"Sub Foo()
    Dim bar(0 To 1) As {variableType}
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Variant")]
        [TestCase("Long")]
        public void VariableTypeNotDeclared_TypedArray_ReDim_DoesNotReturnResult(string variableType)
        {
            var inputCode =
                $@"Sub Foo()
    ReDim bar(0 To 1) As {variableType}
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new VariableTypeNotDeclaredInspection(null);

            Assert.AreEqual(nameof(VariableTypeNotDeclaredInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new VariableTypeNotDeclaredInspection(state);
        }
    }
}

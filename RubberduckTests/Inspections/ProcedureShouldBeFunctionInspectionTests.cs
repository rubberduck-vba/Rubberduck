using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ProcedureShouldBeFunctionInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_ReturnsResult()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef foo As Boolean)
    foo = 42
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_ReturnsResult_UsedByValAfterConditionalAssignment()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef bar As Boolean, ByVal baz As Boolean)
    If baz Then
        bar = 42
    End If
    Goo bar
End Sub

Private Sub Goo(ByVal arg As Boolean)
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_ReturnsResult_UsedByValBeforeAssignment()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef bar As Boolean)
    Goo bar
    bar = 42
End Sub

Private Sub Goo(ByVal arg As Boolean)
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_ReturnsResult_UsedByRef()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef bar As Boolean)
    Goo bar, True
End Sub

Private Sub Goo(ByRef arg1 As Boolean, ByRef arg2 As Boolean)
    arg1 = arg2
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnResult_UsedInExpression()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef bar As Boolean)
    Goo Not bar, True
End Sub

Private Sub Goo(ByRef arg1 As Boolean, ByRef arg2 As Boolean)
    arg1 = arg2
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_ReturnsResult_UsedByRefWithArgumentUsage()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef bar As Boolean)
    Goo bar, True
End Sub

Private Sub Goo(ByRef arg1 As Boolean, ByRef arg2 As Boolean)
    Dim baz As Variant
    baz = arg1
    arg1 = arg2
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_ReturnsResult_MultipleSubs()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef foo As Boolean)
    foo = True
End Sub

Private Sub Goo(ByRef foo As Integer)
    foo = 42
End Sub";

            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnResult_Function()
        {
            const string inputCode =
                @"Private Function Foo(ByRef bar As Integer) As Integer
    Foo = bar
End Function";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnResult_SingleByValParam()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal foo As Integer)
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnsResult_MultipleByValParams()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal foo As Integer, ByVal goo As Integer)
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnsResult_MultipleByRefParams()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef foo As Integer, ByRef goo As Integer)
    foo = 42
    goo = 42
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_ReturnsResult_MultipleParamsOneByRef()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal foo As Integer, ByRef goo As Integer, ByVal hoo As Variant)
    foo = 42
    goo = 42
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnResult_InterfaceImplementation()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByRef a As Integer)
    a = 42
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByRef a As Integer)
End Sub";
            var modules = new (string, string, ComponentType)[] 
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_ReturnsResult_Object()
        {
            const string inputCode1 =
                @"Public bar As Variant";
            const string inputCode2 =
                @"Private Sub DoSomething(ByRef a As Class1)
    Set a = New Class1
End Sub";
            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule)
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnResult_ObjectMember()
        {
            const string inputCode1 =
                @"Public bar As Variant";
            const string inputCode2 =
                @"Private Sub DoSomething(ByRef a As Class1)
    Set a.bar = New Class1
End Sub";
            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_DoesNotReturnResult_EventImplementation()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";
            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
    arg1 = 42
End Sub";

            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule)
            };
            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ProcedureShouldBeFunction_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub Foo(ByRef foo As Boolean)
    foo = 42
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ProcedureCanBeWrittenAsFunctionInspection(null);

            Assert.AreEqual(nameof(ProcedureCanBeWrittenAsFunctionInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ProcedureCanBeWrittenAsFunctionInspection(state);
        }

    }
}

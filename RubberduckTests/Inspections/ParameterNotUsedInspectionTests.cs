using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ParameterNotUsedInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_ReturnsResult()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_ReturnsResult_MultipleSubs()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer)
End Sub

Private Sub Goo(ByVal arg1 as Integer)
End Sub";

            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterUsed_DoesNotReturnResult()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer)
    arg1 = 9
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        //See issue #5336 at https://github.com/rubberduck-vba/Rubberduck/issues/5336
        public void ParameterWithBuiltInTypeNameUsed_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Function Foo(Object As Object) As Object
Set Foo = Object
End Function";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_ReturnsResult_SomeParamsUsed()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer, ByVal arg2 as String)
    arg1 = 9
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        //See issue #4496 at https://github.com/rubberduck-vba/Rubberduck/issues/4496
        public void ParameterNotUsed_RecursiveDefaultMemberAccess_ReturnsNoResult()
        {
            const string inputCode = 
                @"Public Sub Test(rst As ADODB.Recordset)
    Debug.Print rst(""Field"")
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("Module1", inputCode, ComponentType.StandardModule),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules, ReferenceLibrary.AdoDb).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_InterfaceWithImplementation_ReturnsResultForInterface()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
           {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_InterfaceWithImplementation_SomeUseParameter_DoesNotReturnResult()
        {
            const string inputCode1 =
                @"Public Sub DoSomething(ByVal a As Integer)
End Sub";
            const string inputCode2 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
End Sub";
            const string inputCode3 =
                @"Implements IClass1

Private Sub IClass1_DoSomething(ByVal a As Integer)
    Dim bar As Variant
    bar = a
End Sub
";

            var modules = new (string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_InterfaceWithoutImplementation_DoesNotReturnResult()
        {
            const string inputCode1 =
                @"'@Interface
Public Sub DoSomething(ByVal a As Integer)
End Sub";

            var modules = new (string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_EventMemberWithHandlers_ResultForEventOnly()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode2, ComponentType.ClassModule),
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_EventMemberWithHandlers_SomeUseParameter_DoesNotReturnResult()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";

            const string inputCode2 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode3 =
                @"Private WithEvents abc As Class1

Private Sub abc_Foo(ByRef arg1 As Integer)
    Dim bar As Variant
    bar = arg1
End Sub";

            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule),
                ("Class2", inputCode2, ComponentType.ClassModule),
                ("Class3", inputCode3, ComponentType.ClassModule),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_EventMemberWithoutHandlers_DoesNotReturnResult()
        {
            const string inputCode1 =
                @"Public Event Foo(ByRef arg1 As Integer)";

            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_LibraryFunction_DoesNotReturnResult()
        {
            const string inputCode1 =
                @"Public Declare Function MyLibFunction Lib ""MyLib"" (arg1 As Integer) As Integer";

            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_LibraryProcedure_DoesNotReturnResult()
        {
            const string inputCode1 =
                @"Public Declare Sub MyLibProcedure Lib ""MyLib"" (arg1 As Integer)";

            var modules = new (string, string, ComponentType)[]
            {
                ("Class1", inputCode1, ComponentType.ClassModule)
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore ParameterNotUsed
Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_AmbiguousName_DoesNotReturnResult()
        {
            const string inputCode = @"
Public Property Get Item()
    Item = 12
End Property

Public Property Let Item(ByVal Item As Variant)
    DoSomething Item
End Property

Private Sub DoSomething(ByVal value As Variant)
    Msgbox(value)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new ParameterNotUsedInspection(null);

            Assert.AreEqual(nameof(ParameterNotUsedInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ParameterNotUsedInspection(state);
        }
    }
}

using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

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

            Assert.AreEqual(0, InspectionResultsForModules(modules, "ADODB").Count());
        }

        [Test]
        [Category("Inspections")]
        public void ParameterNotUsed_ReturnsResult_InterfaceImplementation()
        {
            //Input
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
           };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
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

using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ConstantNotUsedInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void ConstantNotUsed_ReturnsResult_Local()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Public")]
        [TestCase("Private")]
        public void ConstantUsed_ReturnsResult_Module(string scopeIdentifier)
        {
            var inputCode =
$@"
    {scopeIdentifier} Const Bar As Integer = 42
";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ConstantNotUsed_ReturnsResult_Module_Exposed_Private()
        {
            var inputCode =
$@"
Attribute VB_Exposed = True

    Private Const Bar As Integer = 42
";
            Assert.AreEqual(1, InspectionResultsForModules(("Class1", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void VariableNotUsed_DoesNotReturnResult_Module_Exposed_Public()
        {
            var inputCode =
$@"
Attribute VB_Exposed = True

    Public Const Bar As Integer = 42
";
            Assert.AreEqual(0, InspectionResultsForModules(("Class1", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ConstantNotUsed_ReturnsResult_MultipleConsts()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
    Const const2 As String = ""test""
End Sub";
            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ConstantNotUsed_ReturnsResult_UnusedConstant()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
    Goo const1

    Const const2 As String = ""test""
End Sub

Public Sub Goo(ByVal arg1 As Integer)
End Sub";
            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ConstantNotUsed_UsedConstant_DoesNotReturnResult_Local()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
    Goo const1
End Sub

Public Sub Goo(ByVal arg1 As Integer)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Public")]
        [TestCase("Private")]
        public void ConstantNotUsed_UsedConstant_DoesNotReturnResult_Module(string scopeIdentifier)
        {
            var inputCode =
                $@"
{scopeIdentifier} Const Bar As Integer = 42

Public Sub Foo()
    Goo Bar
End Sub

Public Sub Goo(ByVal arg1 As Integer)
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        // See issue #6042 at https://github.com/rubberduck-vba/Rubberduck/issues/6042
        public void ConstantNotUsed_DoesNotReturnResult_UsedOnlyInArrayUpperBound()
        {
            var inputCode =
                $@"
Sub Test1()
    Const MY_CONST As Byte = 5
    Dim tmpArr(1 To MY_CONST) As Variant           
    Dim tmpArr(1 To MY_CONST)                      
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ConstantNotUsed_DoesNotReturnResult_UsedOnlyInArrayLowerBound()
        {
            var inputCode =
                $@"
Sub Test1()
    Const MY_CONST As Byte = 5
    Dim tmpArr(MY_CONST To 1) As Variant           
    Dim tmpArr(MY_CONST To 1)                      
End Sub
";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        //See issue #4994 at https://github.com/rubberduck-vba/Rubberduck/issues/4994
        public void ConstantNotUsed_ConstantOnlyUsedInMidStatement_DoesNotReturnResult()
        {
            const string inputCode =
                @"Function UnAccent(ByVal inputString As String) As String
          Dim Index As Long, Position As Long
         Const ACCENTED_CHARS As String = ""äéöûü¿¡¬√ƒ≈«»… ÀÃÕŒœ–—“”‘’÷Ÿ⁄€‹›‡·‚„‰ÂÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˘˙˚¸˝ˇ¯ÿüúå""
         Const ANSICHARACTERS As String = ""SZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyyoOYoO""
   For Index = 1 To Len(inputString)
     Position = InStr(ACCENTED_CHARS, Mid$(inputString, Index, 1))
     If Position Then Mid$(inputString, Index) = Mid$(ANSICHARACTERS, Position, 1)
    Next
     UnAccent = inputString
End Function";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ConstantNotUsed_IgnoreModule_All_YieldsNoResult()
        {
            const string inputCode =
                @"'@IgnoreModule

Public Const Bar As Integer = 42

Public Sub Foo()
    Const const1 As Integer = 9
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ConstantNotUsed_IgnoreModule_AnnotationName_YieldsNoResult()
        {
            const string inputCode =
                @"'@IgnoreModule ConstantNotUsed

Public Sub Foo()
    Const const1 As Integer = 9
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void ConstantNotUsed_IgnoreModule_OtherAnnotationName_YieldsResults()
        {
            const string inputCode =
                @"'@IgnoreModule VariableNotUsed

Public Sub Foo()
    Const const1 As Integer = 9
End Sub";
            Assert.IsTrue(InspectionResultsForStandardModule(inputCode).Any());
        }

        [Test]
        [Category("Inspections")]
        public void ConstantNotUsed_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"Public Sub Foo()
    '@Ignore ConstantNotUsed
    Const const1 As Integer = 9
End Sub";
            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ConstantNotUsedInspection";
            var inspection = new ConstantNotUsedInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ConstantNotUsedInspection(state);
        }
    }
}

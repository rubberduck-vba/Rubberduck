using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ObjectVariableNotSetInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenIndexerObjectAccess_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub DoSomething()
    Dim target As Object
    Set target = CreateObject(""Scripting.Dictionary"")
    target(""foo"") = 42
End Sub
";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenIndexerObjectAccess_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub DoSomething()
    Dim target As Object
    target = CreateObject(""Scripting.Dictionary"")
    target(""foo"") = 42
End Sub
";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenStringVariable_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub Workbook_Open()
    
    Dim target As String
    target = Range(""A1"")
    
    target.Value = ""all good""

End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedObject_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant)
'whoCares is a LExprContext and is a known interesting declaration
    Dim target As Collection
    Set target = new Collection
    testParam = target             
    testParam.Add 100
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedNewObject_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant)
'is a NewExprContext
    testParam = new Collection     
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedRange_ReturnsResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant)
'Range(""A1:C1"") is a LExprContext but is not a known interesting declaration
    testParam = Range(""A1:C1"")    
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedDeclaredRange_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant, target As Range)
'target is a LExprContext and is a known interesting declaration
    testParam = target
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedDeclaredVariant_ReturnsNoResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub TestSub(ByRef testParam As Variant, target As Range)
'target is a LExprContext and is a known interesting declaration
    testParam = target
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenVariantVariableAssignedBaseType_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub Workbook_Open()
    Dim target As Variant
    target = ""A1""     'is a LiteralExprContext
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenObjectVariableNotSet_ReturnsResult()
        {
            var expectResultCount = 1;
            var input =
@"
Private Sub Workbook_Open()
    
    Dim target As Range
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenObjectVariableNotSet_Ignored_DoesNotReturnResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub Workbook_Open()
    
    Dim target As Range
'@Ignore ObjectVariableNotSet
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_GivenSetObjectVariable_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub Workbook_Open()
    
    Dim target As Range
    Set target = Range(""A1"")
    
    target.Value = ""All good""

End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2266
        [TestMethod]
        [DeploymentItem(@"Testfiles\")]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_FunctionReturnsArrayOfType_ReturnsNoResult()
        {
            var expectedResultCount = 0;
            var input =
@"
Private Function GetSomeDictionaries() As Dictionary()
    Dim temp(0 To 1) As Worksheet
    Set temp(0) = New Dictionary
    Set temp(1) = New Dictionary
    GetSomeDictionaries = temp
End Function";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Codez", ComponentType.StandardModule, input)
                .AddReference("Scripting", "", 1, 0, true)
                .Build();

            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.State.AddTestLibrary("Scripting.1.0.xml");

            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new ObjectVariableNotSetInspection(parser.State);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(expectedResultCount, inspectionResults.Count());

        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_IgnoreQuickFixWorks()
        {
            var inputCode =
            @"
Private Sub Workbook_Open()
    
    Dim target As Range
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";
            var expectedCode =
            @"
Private Sub Workbook_Open()
    
    Dim target As Range
'@Ignore ObjectVariableNotSet
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObjectVariableNotSetInspection(state);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_ForFunctionAssignment_ReturnsResult()
        {
            var expectedResultCount = 2;
            var input =
@"
Private Function CombineRanges(ByVal source As Range, ByVal toCombine As Range) As Range
    If source Is Nothing Then
        CombineRanges = toCombine 'no inspection result (but there should be one!)
    Else
        CombineRanges = Union(source, toCombine) 'no inspection result (but there should be one!)
    End If
End Function";
            var expectedCode =
            @"
Private Function CombineRanges(ByVal source As Range, ByVal toCombine As Range) As Range
    If source Is Nothing Then
        Set CombineRanges = toCombine 'no inspection result (but there should be one!)
    Else
        Set CombineRanges = Union(source, toCombine) 'no inspection result (but there should be one!)
    End If
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObjectVariableNotSetInspection(state);
            var inspectionResults = inspection.GetInspectionResults().ToList();

            Assert.AreEqual(expectedResultCount, inspectionResults.Count);
            var fix = new UseSetKeywordForObjectAssignmentQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_ForPropertyGetAssignment_ReturnsResults()
        {
            var expectedResultCount = 1;
            var input = @"
Private example As MyObject
Public Property Get Example() As MyObject
    Example = example
End Property
";
            var expectedCode =
            @"
Private example As MyObject
Public Property Get Example() As MyObject
    Set Example = example
End Property
";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(input, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObjectVariableNotSetInspection(state);
            var inspectionResults = inspection.GetInspectionResults().ToList();

            Assert.AreEqual(expectedResultCount, inspectionResults.Count);
            var fix = new UseSetKeywordForObjectAssignmentQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_LongPtrVariable_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub TestLongPtr()
    Dim handle as LongPtr
    handle = 123456
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_NoTypeSpecified_ReturnsResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub TestLongPtr()
    Dim handle as LongPtr
    handle = 123456
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_SelfAssigned_ReturnsNoResult()
        {
            var expectResultCount = 0;
            var input =
@"
Private Sub TestSelfAssigned()
    Dim arg1 As new Collection
    arg1.Add 7
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_EnumVariable_ReturnsNoResult()
        {

            var expectResultCount = 0;
            var input =
@"
Enum TestEnum
    EnumOne
    EnumTwo
    EnumThree
End Enum

Private Sub TestEnum()
    Dim enumVariable As TestEnum
    enumVariable = EnumThree
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObjectVariableNotSet_FunctionReturnNotSet_ReturnsResult()
        {

            var expectResultCount = 0;
            var input =
@"
Enum TestEnum
    EnumOne
    EnumTwo
    EnumThree
End Enum

Private Sub TestEnum()
    Dim enumVariable As TestEnum
    enumVariable = EnumThree
End Sub";
            AssertInputCodeYieldsExpectedInspectionResultCount(input, expectResultCount);
        }

        private void AssertInputCodeYieldsExpectedInspectionResultCount(string inputCode, int expected)
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObjectVariableNotSetInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(expected, inspectionResults.Count());
        }
    }
}


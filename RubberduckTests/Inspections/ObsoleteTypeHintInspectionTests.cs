using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ObsoleteTypeHintInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_FieldWithLongTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo&";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_FieldWithIntegerTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo%";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_FieldWithDoubleTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo#";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_FieldWithSingleTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo!";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_FieldWithDecimalTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo@";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_FieldWithStringTypeHintReturnsResult()
        {
            const string inputCode =
@"Public Foo$";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_FunctionReturnsResult()
        {
            const string inputCode =
@"Public Function Foo$(ByVal bar As Boolean)
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_PropertyGetReturnsResult()
        {
            const string inputCode =
@"Public Property Get Foo$(ByVal bar As Boolean)
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_ParameterReturnsResult()
        {
            const string inputCode =
@"Public Function Foo(ByVal bar$) As Boolean
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_VariableReturnsResult()
        {
            const string inputCode =
@"Public Function Foo() As Boolean
    Dim buzz$
    Foo = True
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_StringValueDoesNotReturnsResult()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim bar As String
    bar = ""Public baz$""
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_FieldsReturnMultipleResults()
        {
            const string inputCode =
@"Public Foo$
Public Bar$";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore ObsoleteTypeHint
Public Function Foo$(ByVal bar As Boolean)
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_LongTypeHint()
        {
            const string inputCode =
@"Public Foo&";

            const string expectedCode =
@"Public Foo As Long";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var fix = new RemoveTypeHintsQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_IntegerTypeHint()
        {
            const string inputCode =
@"Public Foo%";

            const string expectedCode =
@"Public Foo As Integer";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var fix = new RemoveTypeHintsQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_DoubleTypeHint()
        {
            const string inputCode =
@"Public Foo#";

            const string expectedCode =
@"Public Foo As Double";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var fix = new RemoveTypeHintsQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_SingleTypeHint()
        {
            const string inputCode =
@"Public Foo!";

            const string expectedCode =
@"Public Foo As Single";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var fix = new RemoveTypeHintsQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_DecimalTypeHint()
        {
            const string inputCode =
@"Public Foo@";

            const string expectedCode =
@"Public Foo As Decimal";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var fix = new RemoveTypeHintsQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_StringTypeHint()
        {
            const string inputCode =
@"Public Foo$";

            const string expectedCode =
@"Public Foo As String";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var fix = new RemoveTypeHintsQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_QuickFixWorks_Function_StringTypeHint()
        {
            const string inputCode =
@"Public Function Foo$(ByVal fizz As Integer)
    Foo = ""test""
End Function";

            const string expectedCode =
@"Public Function Foo(ByVal fizz As Integer) As String
    Foo = ""test""
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var fix = new RemoveTypeHintsQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_QuickFixWorks_PropertyGet_StringTypeHint()
        {
            const string inputCode =
@"Public Property Get Foo$(ByVal fizz As Integer)
    Foo = ""test""
End Property";

            const string expectedCode =
@"Public Property Get Foo(ByVal fizz As Integer) As String
    Foo = ""test""
End Property";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();
            
            var fix = new RemoveTypeHintsQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_QuickFixWorks_Parameter_StringTypeHint()
        {
            const string inputCode =
@"Public Sub Foo(ByVal fizz$)
    Foo = ""test""
End Sub";

            const string expectedCode =
@"Public Sub Foo(ByVal fizz As String)
    Foo = ""test""
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var fix = new RemoveTypeHintsQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_QuickFixWorks_Variable_StringTypeHint()
        {
            const string inputCode =
@"Public Sub Foo()
    Dim buzz$
End Sub";

            const string expectedCode =
@"Public Sub Foo()
    Dim buzz As String
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var fix = new RemoveTypeHintsQuickFix(state);
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteTypeHint_IgnoreQuickFixWorks()
        {
            const string inputCode =
@"Public Function Foo$(ByVal fizz As Integer)
    Foo = ""test""
End Function";

            const string expectedCode =
@"'@Ignore ObsoleteTypeHint
Public Function Foo$(ByVal fizz As Integer)
    Foo = ""test""
End Function";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteTypeHintInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            var fix = new IgnoreOnceQuickFix(state, new[] {inspection});
            foreach (var result in inspectionResults)
            {
                fix.Fix(result);
            }

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ObsoleteTypeHintInspection(null);
            Assert.AreEqual(CodeInspectionType.LanguageOpportunities, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteTypeHintInspection";
            var inspection = new ObsoleteTypeHintInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}

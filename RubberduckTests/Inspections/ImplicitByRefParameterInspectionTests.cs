using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ImplicitByRefParameterInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo(arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_ReturnsResult_MultipleParams()
        {
            const string inputCode =
@"Sub Foo(arg1 As Integer, arg2 As Date)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_DoesNotReturnResult_ByRef()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_DoesNotReturnResult_ByVal()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_DoesNotReturnResult_ParamArray()
        {
            const string inputCode =
@"Sub Foo(ParamArray arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_ReturnsResult_SomePassedByRefImplicitely()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 As Integer, arg2 As Date)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_ReturnsResult_InterfaceImplementation()
        {
            //Input
            const string inputCode1 =
@"Sub Foo(arg1 As Integer)
End Sub";
            const string inputCode2 =
@"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore ImplicitByRefParameter
Sub Foo(arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_QuickFixWorks_PassByRef()
        {
            const string inputCode =
@"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedCode =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ChangeParameterByRefByValQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_IgnoredQuickFixWorks()
        {
            const string inputCode =
@"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedCode =
@"'@Ignore ImplicitByRefParameter
Sub Foo(arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();
            
            new IgnoreOnceQuickFix(state, new[] {inspection}).Fix(inspectionResults.First());
            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        //http://chat.stackexchange.com/transcript/message/34001991#34001991
        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_QuickFixWorks_OptionalParameter()
        {
            const string inputCode =
@"Sub Foo(Optional arg1 As Integer)
End Sub";

            const string expectedCode =
@"Sub Foo(Optional ByRef arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ChangeParameterByRefByValQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        //http://chat.stackexchange.com/transcript/message/34001991#34001991
        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_QuickFixWorks_Optional_LineContinuations()
        {
            const string inputCode =
@"Sub Foo(Optional _
        bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
@"Sub Foo(Optional _
        ByRef bar _
        As Byte)
    bar = 1
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ChangeParameterByRefByValQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        //http://chat.stackexchange.com/transcript/message/34001991#34001991
        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_QuickFixWorks_LineContinuation()
        {
            const string inputCode =
@"Sub Foo( bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
@"Sub Foo( ByRef bar _
        As Byte)
    bar = 1
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ChangeParameterByRefByValQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        //http://chat.stackexchange.com/transcript/message/34001991#34001991
        [TestMethod]
        [TestCategory("Inspections")]
        public void ImplicitByRefParameter_QuickFixWorks_LineContinuation_FirstLine()
        {
            const string inputCode =
@"Sub Foo( _
        bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
@"Sub Foo( _
        ByRef bar _
        As Byte)
    bar = 1
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ImplicitByRefParameterInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            new ChangeParameterByRefByValQuickFix(state).Fix(inspectionResults.First());

            Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ImplicitByRefParameterInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitByRefParameterInspection";
            var inspection = new ImplicitByRefParameterInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}

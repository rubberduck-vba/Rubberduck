using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class ObsoleteByRefModifierInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_ReturnsResult()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_ReturnsResult_MultipleParameters()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As Integer, ByRef arg2 As Date)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(2, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_DoesNotReturnResult_NoModifier()
        {
            const string inputCode =
@"Sub Foo(arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_DoesNotReturnResult_ByVal()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_ReturnsResult_SomePassedByRef()
        {
            const string inputCode =
@"Sub Foo(ByVal arg1 As Integer, ByRef arg2 As Date)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_ReturnsResult_InterfaceImplementation()
        {
            const string inputCode1 =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode2 =
@"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_ReturnsResult_MultipleInterfaceImplementation()
        {
            const string inputCode1 =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode2 =
@"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode3 =
@"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode3)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
@"'@Ignore ObsoleteByRefModifier
Sub Foo(ByRef arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            Assert.IsFalse(inspectionResults.Any());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_QuickFixWorks_PassByRef()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
@"Sub Foo(arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_IgnoredQuickFixWorks()
        {
            const string inputCode =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
@"'@Ignore ObsoleteByRefModifier
Sub Foo(ByRef arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is IgnoreOnceQuickFix).Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_QuickFixWorks_OptionalParameter()
        {
            const string inputCode =
@"Sub Foo(Optional ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
@"Sub Foo(Optional arg1 As Integer)
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_QuickFixWorks_Optional_LineContinuations()
        {
            const string inputCode =
@"Sub Foo(Optional ByRef _
        bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
@"Sub Foo(Optional _
        bar _
        As Byte)
    bar = 1
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_QuickFixWorks_LineContinuation()
        {
            const string inputCode =
@"Sub Foo( ByRef bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
@"Sub Foo( bar _
        As Byte)
    bar = 1
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_QuickFixWorks_LineContinuation_FirstLine()
        {
            const string inputCode =
@"Sub Foo( _
        ByRef bar _
        As Byte)
    bar = 1
End Sub";

            const string expectedCode =
@"Sub Foo( _
        bar _
        As Byte)
    bar = 1
End Sub";

            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            Assert.AreEqual(expectedCode, component.CodeModule.Content());
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void ObsoleteByRefModifier_QuickFixWorks_InterfaceImplementation()
        {
            const string inputCode1 =
@"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode2 =
@"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode1 =
@"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedCode2 =
@"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            var state = MockParser.CreateAndParse(vbe.Object);

            var inspection = new ObsoleteByRefModifierInspection(state);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.First().Fix();

            var project = vbe.Object.VBProjects[0];
            var interfaceCode = project.VBComponents[0].CodeModule.Content();
            var implementationCode = project.VBComponents[1].CodeModule.Content();

            Assert.AreEqual(expectedCode1, interfaceCode, "Wrong code in interface");
            Assert.AreEqual(expectedCode2, implementationCode, "Wrong code in first implementation");
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new ObsoleteByRefModifierInspection(null);
            Assert.AreEqual(CodeInspectionType.CodeQualityIssues, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ObsoleteByRefModifierInspection";
            var inspection = new ObsoleteByRefModifierInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}

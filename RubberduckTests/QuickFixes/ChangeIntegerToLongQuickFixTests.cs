using System.Linq;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ChangeIntegerToLongQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_Function()
        {
            const string inputCode =
                @"Function Foo() As Integer
End Function";

            const string expectedCode =
                @"Function Foo() As Long
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_FunctionWithTypeHint()
        {
            const string inputCode =
                @"Function Foo%()
End Function";

            const string expectedCode =
                @"Function Foo&()
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGet()
        {
            const string inputCode =
                @"Property Get Foo() As Integer
End Property";

            const string expectedCode =
                @"Property Get Foo() As Long
End Property";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGetWithTypeHint()
        {
            const string inputCode =
                @"Property Get Foo%()
End Property";

            const string expectedCode =
                @"Property Get Foo&()
End Property";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_Parameter()
        {
            const string inputCode =
                @"Sub Foo(arg As Integer)
End Sub";

            const string expectedCode =
                @"Sub Foo(arg As Long)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterWithTypeHint()
        {
            const string inputCode =
                @"Sub Foo(arg%)
End Sub";

            const string expectedCode =
                @"Sub Foo(arg&)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_Variable()
        {
            const string inputCode =
                @"Sub Foo()
    Dim v As Integer
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim v As Long
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_VariableWithTypeHint()
        {
            const string inputCode =
                @"Sub Foo()
    Dim v%
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim v&
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_Constant()
        {
            const string inputCode =
                @"Sub Foo()
    Const c As Integer = 0
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Const c As Long = 0
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ConstantWithTypeHint()
        {
            const string inputCode =
                @"Sub Foo()
    Const c% = 0
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Const c& = 0
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_UserDefinedTypeReservedNameMember()
        {
            const string inputCode =
                @"Type T
    i as Integer
End Type";

            const string expectedCode =
                @"Type T
    i as Long
End Type";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_UserDefinedTypeUntypedNameMember()
        {
            const string inputCode =
                @"Type T
    i() as Integer
End Type";

            const string expectedCode =
                @"Type T
    i() as Long
End Type";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_FunctionInterfaceImplementation()
        {
            const string inputCode1 =
                @"Function Foo() As Integer
End Function";

            const string inputCode2 =
                @"Implements IClass1

Function IClass1_Foo() As Integer
End Function";

            const string expectedCode1 =
                @"Function Foo() As Long
End Function";

            const string expectedCode2 =
                @"Implements IClass1

Function IClass1_Foo() As Long
End Function";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_FunctionInterfaceImplementationWithTypeHints()
        {
            const string inputCode1 =
                @"Function Foo%()
End Function";

            const string inputCode2 =
                @"Implements IClass1

Function IClass1_Foo%()
End Function";

            const string expectedCode1 =
                @"Function Foo&()
End Function";

            const string expectedCode2 =
                @"Implements IClass1

Function IClass1_Foo&()
End Function";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_FunctionInterfaceImplementationWithInterfaceTypeHint()
        {
            const string inputCode1 =
                @"Function Foo%()
End Function";

            const string inputCode2 =
                @"Implements IClass1

Function IClass1_Foo() As Integer
End Function";

            const string expectedCode1 =
                @"Function Foo&()
End Function";

            const string expectedCode2 =
                @"Implements IClass1

Function IClass1_Foo() As Long
End Function";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_FunctionInterfaceImplementationWithImplementationTypeHint()
        {
            const string inputCode1 =
                @"Function Foo() As Integer
End Function";

            const string inputCode2 =
                @"Implements IClass1

Function IClass1_Foo%()
End Function";

            const string expectedCode1 =
                @"Function Foo() As Long
End Function";

            const string expectedCode2 =
                @"Implements IClass1

Function IClass1_Foo&()
End Function";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGetInterfaceImplementation()
        {
            const string inputCode1 =
                @"Property Get Foo() As Integer
End Property";

            const string inputCode2 =
                @"Implements IClass1

Property Get IClass1_Foo() As Integer
End Property";

            const string expectedCode1 =
                @"Property Get Foo() As Long
End Property";

            const string expectedCode2 =
                @"Implements IClass1

Property Get IClass1_Foo() As Long
End Property";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGetInterfaceImplementationWithTypeHints()
        {
            const string inputCode1 =
                @"Property Get Foo%()
End Property";

            const string inputCode2 =
                @"Implements IClass1

Property Get IClass1_Foo%()
End Property";

            const string expectedCode1 =
                @"Property Get Foo&()
End Property";

            const string expectedCode2 =
                @"Implements IClass1

Property Get IClass1_Foo&()
End Property";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGetInterfaceImplementationWithInterfaceTypeHint()
        {
            const string inputCode1 =
                @"Property Get Foo%()
End Property";

            const string inputCode2 =
                @"Implements IClass1

Property Get IClass1_Foo() As Integer
End Property";

            const string expectedCode1 =
                @"Property Get Foo&()
End Property";

            const string expectedCode2 =
                @"Implements IClass1

Property Get IClass1_Foo() As Long
End Property";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_PropertyGetInterfaceImplementationWithImplementationTypeHint()
        {
            const string inputCode1 =
                @"Property Get Foo() As Integer
End Property";

            const string inputCode2 =
                @"Implements IClass1

Property Get IClass1_Foo%()
End Property";

            const string expectedCode1 =
                @"Property Get Foo() As Long
End Property";

            const string expectedCode2 =
                @"Implements IClass1

Property Get IClass1_Foo&()
End Property";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterInterfaceImplementationWithTypeHints()
        {
            const string inputCode1 =
                @"Sub Foo(arg1%)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1%)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(arg1&)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1&)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterInterfaceImplementationWithInterfaceTypeHint()
        {
            const string inputCode1 =
                @"Sub Foo(arg1%)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(arg1&)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Long)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterInterfaceImplementationWithImplementationTypeHint()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1%)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(arg1 As Long)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1&)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterInterfaceImplementation()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(arg1 As Long)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Long)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_ParameterInterfaceImplementationWithDifferentName()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg2 As Integer)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(arg1 As Long)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg2 As Long)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(inspectionResults.First());

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void IntegerDataType_QuickFixWorks_MultipleParameterInterfaceImplementation()
        {
            const string inputCode1 =
                @"Sub Foo(arg1 As Integer, arg2 as Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Integer, arg2 as Integer)
End Sub";

            const string expectedCode1 =
                @"Sub Foo(arg1 As Long, arg2 as Integer)
End Sub";

            const string expectedCode2 =
                @"Implements IClass1

Sub IClass1_Foo(arg1 As Long, arg2 as Integer)
End Sub";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode2)
                .MockVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new IntegerDataTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new ChangeIntegerToLongQuickFix(state).Fix(
                    inspectionResults.First(
                        result =>
                            ((VBAParser.ArgContext)result.Context).unrestrictedIdentifier()
                            .identifier()
                            .untypedIdentifier()
                            .identifierValue()
                            .GetText() == "arg1"));

                var project = vbe.Object.VBProjects[0];
                var interfaceComponent = project.VBComponents[0];
                var implementationComponent = project.VBComponents[1];

                Assert.AreEqual(expectedCode1, state.GetRewriter(interfaceComponent).GetText(), "Wrong code in interface");
                Assert.AreEqual(expectedCode2, state.GetRewriter(implementationComponent).GetText(), "Wrong code in implementation");
            }
        }

    }
}

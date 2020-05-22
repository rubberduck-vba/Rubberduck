using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using RubberduckTests.Mocks;

namespace RubberduckTests.Grammar
{
    [TestFixture]
    public class AttributesTests
    {
        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ModuleAttributeAtStartGetsRecognized()
        {
            const string inputCode =
                @"Attribute VB_Description = ""Whatever""

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName; 
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = (moduleName, DeclarationType.ProceduralModule);
                var expectedAttributeName = "VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ModuleAttributeBetweenModuleVariableDeclarationsGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long
Attribute VB_Description = ""Whatever""
Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = (moduleName, DeclarationType.ProceduralModule);
                var expectedAttributeName = "VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ModuleAttributeAfterModuleDeclarationsGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String
Attribute VB_Description = ""Whatever""
Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = (moduleName, DeclarationType.ProceduralModule);
                var expectedAttributeName = "VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ModuleAttributeInModuleBodyElementGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
Attribute VB_Description = ""Whatever""
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = (moduleName, DeclarationType.ProceduralModule);
                var expectedAttributeName = "VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ModuleAttributeBetweenModuleBodyElementsGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Attribute VB_Description = ""Whatever""

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = (moduleName, DeclarationType.ProceduralModule);
                var expectedAttributeName = "VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ModuleAttributeAtEndGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

Attribute VB_Description = ""Whatever""";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = (moduleName, DeclarationType.ProceduralModule);
                var expectedAttributeName = "VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void VariableAttributeBeforeVariableDeclarationDoesNotGetRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Attribute baz.VB_VarDescription = ""Whatever""

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleAttributes = state.GetModuleAttributes(module);
                Assert.IsFalse(moduleAttributes.Any());
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void VariableAttributeBetweenVariableDeclarationsAfterTheDeclarationGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Attribute bar.VB_VarDescription = ""Whatever""

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("bar", DeclarationType.Variable);
                var expectedAttributeName = "bar.VB_VarDescription";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void VariableAttributeBetweenVariableDeclarationsAndModuleBodyGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Attribute bar.VB_VarDescription = ""Whatever""
Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("bar", DeclarationType.Variable);
                var expectedAttributeName = "bar.VB_VarDescription";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void VariableAttributeInModuleBodyElementGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
Attribute bar.VB_VarDescription = ""Whatever""
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("bar", DeclarationType.Variable);
                var expectedAttributeName = "bar.VB_VarDescription";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void VariableAttributeBetweenModuleBodyElementsGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Attribute bar.VB_VarDescription = ""Whatever""

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("bar", DeclarationType.Variable);
                var expectedAttributeName = "bar.VB_VarDescription";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void VariableAttributeAtTheEndGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

Attribute bar.VB_VarDescription = ""Whatever""";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("bar", DeclarationType.Variable);
                var expectedAttributeName = "bar.VB_VarDescription";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void VariableAttributeForVariableInADeclarationListGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer
Attribute bar.VB_VarDescription = ""Whatever""
Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("bar", DeclarationType.Variable);
                var expectedAttributeName = "bar.VB_VarDescription";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ProcedureAttributeInsideTheProcedureGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
Attribute Foo.VB_Description = ""Whatever""
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("Foo", DeclarationType.Procedure);
                var expectedAttributeName = "Foo.VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ProcedureAttributeRightAfterTheProcedureGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub
Attribute Foo.VB_Description = ""Whatever""

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("Foo", DeclarationType.Procedure);
                var expectedAttributeName = "Foo.VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ProcedureAttributeAfterTheProcedureDoesNotGetRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Attribute Foo.VB_Description = ""Whatever""

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleAttributes = state.GetModuleAttributes(module);
                Assert.IsFalse(moduleAttributes.Any());
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ProcedureAttributeBeforeTheProcedureDoesNotGetRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Attribute Foo.VB_Description = ""Whatever""
Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleAttributes = state.GetModuleAttributes(module);
                Assert.IsFalse(moduleAttributes.Any());
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void FunctionAttributeInsideTheFunctionGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Attribute Woo.VB_Description = ""Whatever""
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("Woo", DeclarationType.Function);
                var expectedAttributeName = "Woo.VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void FunctionAttributeRightAfterTheFunctionGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function
Attribute Woo.VB_Description = ""Whatever""

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("Woo", DeclarationType.Function);
                var expectedAttributeName = "Woo.VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void FunctionAttributeAfterTheFunctionDoesNotGetRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Attribute Woo.VB_Description = ""Whatever""

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleAttributes = state.GetModuleAttributes(module);
                Assert.IsFalse(moduleAttributes.Any());
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void FunctionAttributeBeforeTheFunctionDoesNotGetRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Attribute Woo.VB_Description = ""Whatever""
Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleAttributes = state.GetModuleAttributes(module);
                Assert.IsFalse(moduleAttributes.Any());
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertyGetAttributeInsideThePropertyGetGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
Attribute FooBar.VB_Description = ""Whatever""
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("FooBar", DeclarationType.PropertyGet);
                var expectedAttributeName = "FooBar.VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertyGetAttributeRightAfterThePropertyGetGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property
Attribute FooBar.VB_Description = ""Whatever""

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("FooBar", DeclarationType.PropertyGet);
                var expectedAttributeName = "FooBar.VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertyGetAttributeAfterThePropertyGetDoesNotGetRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Attribute FooBar.VB_Description = ""Whatever""

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleAttributes = state.GetModuleAttributes(module);
                Assert.IsFalse(moduleAttributes.Any());
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertyGetAttributeBeforeThePropertyGetDoesNotGetRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Attribute Foo.VB_Description = ""Whatever""
Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Attribute FooBar.VB_Description = ""Whatever""
Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleAttributes = state.GetModuleAttributes(module);
                Assert.IsFalse(moduleAttributes.Any());
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertyLetAttributeInsideThePropertyLetGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
Attribute FooBar.VB_Description = ""Whatever""
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("FooBar", DeclarationType.PropertyLet);
                var expectedAttributeName = "FooBar.VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertyLetAttributeRightAfterThePropertyLetGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property
Attribute FooBar.VB_Description = ""Whatever""

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("FooBar", DeclarationType.PropertyLet);
                var expectedAttributeName = "FooBar.VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertyLetAttributeAfterThePropertyLetDoesNotGetRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Attribute Foo.VB_Description = ""Whatever""

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Attribute FooBar.VB_Description = ""Whatever""

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleAttributes = state.GetModuleAttributes(module);
                Assert.IsFalse(moduleAttributes.Any());
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertyLetAttributeBeforeThePropertyLetDoesNotGetRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Attribute Foo.VB_Description = ""Whatever""
Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As LongPropertyLet
    FooBar = 0
End Property

Attribute FooBar.VB_Description = ""Whatever""
Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleAttributes = state.GetModuleAttributes(module);
                Assert.IsFalse(moduleAttributes.Any());
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertySetAttributeInsideThePropertySetGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
Attribute FooBaz.VB_Description = ""Whatever""
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("FooBaz", DeclarationType.PropertySet);
                var expectedAttributeName = "FooBaz.VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertySetAttributeRightAfterThePropertySetGetsRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property
Attribute FooBaz.VB_Description = ""Whatever""

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScopeCount = 1;
                var expectedAttributeScope = ("FooBaz", DeclarationType.PropertySet);
                var expectedAttributeName = "FooBaz.VB_Description";
                var expectedAttributeValue = @"""Whatever""";

                var moduleAttributes = state.GetModuleAttributes(module);
                var moduleAttributesCount = moduleAttributes.Count;
                Assert.AreEqual(expectedAttributeScopeCount, moduleAttributesCount);
                var scopeAttributes = moduleAttributes[expectedAttributeScope];
                var attribute = scopeAttributes.Single();
                var attributeName = attribute.Name;
                var value = attribute.Values.Single();
                Assert.AreEqual(expectedAttributeName, attributeName);
                Assert.AreEqual(expectedAttributeValue, value);
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertySetAttributeAfterThePropertySetDoesNotGetRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

Attribute FooBaz.VB_Description = ""Whatever""
";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleAttributes = state.GetModuleAttributes(module);
                Assert.IsFalse(moduleAttributes.Any());
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertySetAttributeBeforeThePropertySetDoesNotGetRecognized()
        {
            const string inputCode =
                @"

Public bar As Long

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
End Sub

Public Function Woo() As String
    Woo = ""test""
End Function

Public Property Get FooBar() As Long
    FooBar = 0
End Property

Public Property Let FooBar(arg As Long)
End Property

Attribute FooBaz.VB_Description = ""Whatever""
Public Property Set FooBaz(arg As Object)
    Set FooBaz = arg
End Property

";
            var moduleName = "TestModule";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, moduleName, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var moduleAttributes = state.GetModuleAttributes(module);
                Assert.IsFalse(moduleAttributes.Any());
            }
        }
    }
}
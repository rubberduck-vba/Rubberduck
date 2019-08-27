using System.Linq;
using Antlr4.Runtime;
using NUnit.Framework;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.TypeResolvers;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.VBA.Parsing;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.ExpressionResolving
{
    [TestFixture]
    public class SetTypeResolverTests
    {
        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("TestProject.Class1", "TestProject.Class1")]
        [TestCase("Variant", "Variant")]
        [TestCase("Object", "Object")]
        [TestCase("Long", SetTypeResolver.NotAnObject)]
        public void SimpleNameExpression_SetTypeNameTests(string typeName, string expectedSetTypeName)
        {
            const string class1 =
                @"
Private Sub Foo()
End Sub
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {typeName}
    Set cls = cls
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 18);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("TestProject.Class1", "TestProject.Class1")]
        [TestCase("Variant", null)]
        [TestCase("Object", null)]
        [TestCase("Long", null)]
        public void SimpleNameExpression_SetTypeDeclarationTests(string typeName, string expectedNameOfSetTypeDeclaration)
        {
            const string class1 =
                @"
Private Sub Foo()
End Sub
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {typeName}
    Set cls = cls
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 18);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        public void InstanceExpression_SetTypeName_ReturnsNameOfContainingClass()
        {
            const string class1 =
                @"
Private Sub Foo()
    Dim bar As Variant
    Set bar = Me
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 17);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Class1", expressionSelection);
            var expectedSetTypeName = "TestProject.Class1";

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        public void InstanceExpression_SetTypeDeclaration_ReturnsDeclarationOfContainingClass()
        {
            const string class1 =
                @"
Private Sub Foo()
    Dim bar As Variant
    Set bar = Me
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 17);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Class1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();
            var expectedNameOfSetTypeDeclaration = "TestProject.Class1";

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("TestProject.Class1", "TestProject.Class1")]
        [TestCase("Variant", "Variant")]
        [TestCase("Object", "Object")]
        [TestCase("Long", SetTypeResolver.NotAnObject)]
        public void MemberAccessExpression_SetTypeNameTests(string typeName, string expectedSetTypeName)
        {
            var class1 =
                $@"
Public Property Get Foo() As {typeName}
End Property
";

            var module1 =
                @"
Private Sub Bar()
    Dim cls As Class1
    Dim baz as Variant
    Set baz = cls.Foo
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 22);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("TestProject.Class1", "TestProject.Class1")]
        [TestCase("Variant", null)]
        [TestCase("Object", null)]
        [TestCase("Long", null)]
        public void MemberAccessExpression_SetTypeDeclarationTests(string typeName, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                $@"
Public Property Get Foo() As {typeName}
End Property
";

            var module1 =
                @"
Private Sub Bar()
    Dim cls As Class1
    Dim baz As Variant
    Set baz = cls.Foo
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 22);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("TestProject.Class1", "TestProject.Class1")]
        [TestCase("Variant", "Variant")]
        [TestCase("Object", "Object")]
        [TestCase("Long", SetTypeResolver.NotAnObject)]
        public void WithMemberAccessExpression_SetTypeNameTests(string typeName, string expectedSetTypeName)
        {
            var class1 =
                $@"
Public Property Get Foo() As {typeName}
End Property
";

            var module1 =
                @"
Private Sub Bar()
    With New Class1
        Dim baz as Variant
        Set baz = .Foo
    End With
End Sub
";

            var expressionSelection = new Selection(5, 19, 5, 23);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("TestProject.Class1", "TestProject.Class1")]
        [TestCase("Variant", null)]
        [TestCase("Object", null)]
        [TestCase("Long", null)]
        public void WithMemberAccessExpression_SetTypeDeclarationTests(string typeName, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                $@"
Public Property Get Foo() As {typeName}
End Property
";

            var module1 =
                @"
Private Sub Bar()
    With New Class1
        Dim baz as Variant
        Set baz = .Foo
    End With
End Sub
";

            var expressionSelection = new Selection(5, 19, 5, 23);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        public void TypeOfExpression_SetTypeName_ReturnsSetTypeNameOfExpression()
        {
            var class1 =
                @"
Public Property Get Foo() As Class2
End Property
";
            var class2 =
                @"";

            var module1 =
                @"
Private Sub Bar()
    Dim cls as Class1
    Dim baz as Variant
    baz = TypeOf cls.Foo Is TestProject.Class1
End Sub
";

            var expressionSelection = new Selection(5, 11, 5, 25);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Class2", ComponentType.ClassModule, class2)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);
            var expectedSetTypeName = "TestProject.Class2";

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        public void TypeOfExpression_SetTypeDeclaration_ReturnsSetTypeDeclarationOfExpression()
        {
            var class1 =
                @"
Public Property Get Foo() As Class2
End Property
";
            var class2 =
                @"";

            var module1 =
                @"
Private Sub Bar()
    Dim cls as Class1
    Dim baz as Variant
    baz = TypeOf cls.Foo Is TestProject.Class1
End Sub
";

            var expressionSelection = new Selection(5, 11, 5, 25);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Class2", ComponentType.ClassModule, class2)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();
            var expectedNameOfSetTypeDeclaration = "TestProject.Class2";

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("TestProject.Class1", "TestProject.Class1")]
        public void NewExpression_SetTypeNameTests(string typeName, string expectedSetTypeName)
        {
            var class1 =
                @"
Public Property Get Foo() As Variant
End Property
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz as Variant
    Set baz = New {typeName}
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 20);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("TestProject.Class1", "TestProject.Class1")]
        public void NewExpression_SetTypeDeclarationTests(string typeName, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                @"
Public Property Get Foo() As Variant
End Property
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz as Variant
    Set baz = New {typeName}
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 20);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Nothing", "Object")]
        [TestCase("5", SetTypeResolver.NotAnObject)]
        public void LiteralExpression_SetTypeNameTests(string literal, string expectedSetTypeName)
        {
            var class1 =
                @"
Public Property Get Foo() As Variant
End Property
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz as Variant
    Set baz = {literal}
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 16);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Nothing", null)]
        [TestCase("5", null)]
        public void LiteralExpression_SetTypeDeclarationTests(string literal, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                @"
Public Property Get Foo() As Variant
End Property
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz as Variant
    Set baz = {literal}
End Sub
";

            var expressionSelection = new Selection(4, 15, 4, 16);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object")]
        [TestCase("[Object]", "Object")]
        [TestCase("Variant", "Variant")]
        [TestCase("[Variant]", "Variant")]
        [TestCase("Any", "Any")]
        [TestCase("[Any]", "Any")]
        [TestCase("Long", SetTypeResolver.NotAnObject)]
        [TestCase("[Long]", SetTypeResolver.NotAnObject)]
        public void BuiltInTypeExpression_SetTypeNameTests(string builtInType, string expectedSetTypeName)
        {
            var class1 =
                @"
Public Property Get Foo() As Variant
End Property
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz as Variant
    baz = TypeOf baz Is {builtInType}
End Sub
";

            var expressionSelection = new Selection(4, 25, 4, 26);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", null)]
        [TestCase("[Object]", null)]
        [TestCase("Variant", null)]
        [TestCase("[Variant]", null)]
        [TestCase("Any", null)]
        [TestCase("[Any]", null)]
        [TestCase("Long", null)]
        [TestCase("[Long]", null)]
        public void BuiltInTypeExpression_SetTypeDeclarationTests(string builtInType, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                @"
Public Property Get Foo() As Variant
End Property
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz as Variant
    baz = TypeOf baz Is {builtInType}
End Sub
";

            var expressionSelection = new Selection(4, 25, 4, 26);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object", null)]
        [TestCase("Class1", "Object", "Object")]
        [TestCase("Variant", "Variant", null)]
        [TestCase("Class1", "Variant", "Variant")]
        [TestCase("Object", "Class1", null)]
        [TestCase("Variant", "Class1", null)]
        [TestCase("Class1", "Class1", "TestProject.Class1")]
        [TestCase("Class1", "Long", SetTypeResolver.NotAnObject)]
        [TestCase("Object", "Long", null)]
        [TestCase("Variant", "Long", null)]
        public void DictionaryAccessExpression_SetTypeNameTests(string accessedTypeName, string typeName, string expectedSetTypeName)
        {
            var class1 =
                $@"
Public Function Foo(baz As String) As {typeName}
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {accessedTypeName}
    Dim baz As Variant
    Set baz = cls!whatever 
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 27);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object", null)]
        [TestCase("Class1", "Object", null)]
        [TestCase("Variant", "Variant", null)]
        [TestCase("Class1", "Variant", null)]
        [TestCase("Object", "Class1", null)]
        [TestCase("Variant", "Class1", null)]
        [TestCase("Class1", "Class1", "TestProject.Class1")]
        [TestCase("Class1", "Long", null)]
        [TestCase("Object", "Long", null)]
        [TestCase("Variant", "Long", null)]
        public void DictionaryAccessExpression_SetTypeDeclarationTests(string accessedTypeName, string typeName, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                $@"
Public Function Foo(baz As String) As {typeName}
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {accessedTypeName}
    Dim baz As Variant
    Set baz = cls!whatever 
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 27);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object", null)]
        [TestCase("Class1", "Object", "Object")]
        [TestCase("Variant", "Variant", null)]
        [TestCase("Class1", "Variant", "Variant")]
        [TestCase("Object", "Class1", null)]
        [TestCase("Variant", "Class1", null)]
        [TestCase("Class1", "Class1", "TestProject.Class1")]
        [TestCase("Class1", "Long", SetTypeResolver.NotAnObject)]
        [TestCase("Object", "Long", null)]
        [TestCase("Variant", "Long", null)]
        public void WithDictionaryAccessExpression_SetTypeNameTests(string accessedTypeName, string typeName, string expectedSetTypeName)
        {
            var class1 =
                $@"
Public Function Foo(baz As String) As {typeName}
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {accessedTypeName}
    Dim baz As Variant
    With cls 
        Set baz = !whatever
    End With
End Sub
";

            var expressionSelection = new Selection(6, 19, 6, 28);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object", null)]
        [TestCase("Class1", "Object", null)]
        [TestCase("Variant", "Variant", null)]
        [TestCase("Class1", "Variant", null)]
        [TestCase("Object", "Class1", null)]
        [TestCase("Variant", "Class1", null)]
        [TestCase("Class1", "Class1", "TestProject.Class1")]
        [TestCase("Class1", "Long", null)]
        [TestCase("Object", "Long", null)]
        [TestCase("Variant", "Long", null)]
        public void WithDictionaryAccessExpression_SetTypeDeclarationTests(string accessedTypeName, string typeName, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                $@"
Public Function Foo(baz As String) As {typeName}
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {accessedTypeName}
    Dim baz As Variant
    With cls 
        Set baz = !whatever
    End With
End Sub
";

            var expressionSelection = new Selection(6, 19, 6, 28);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object", null)]
        [TestCase("Class1", "Object", "Object")]
        [TestCase("Variant", "Variant", null)]
        [TestCase("Class1", "Variant", "Variant")]
        [TestCase("Object", "Class1", null)]
        [TestCase("Variant", "Class1", null)]
        [TestCase("Class1", "Class1", "TestProject.Class1")]
        [TestCase("Class1", "Long", SetTypeResolver.NotAnObject)]
        [TestCase("Object", "Long", null)]
        [TestCase("Variant", "Long", null)]
        public void DefaultMemberIndexExpression_SetTypeNameTests(string accessedTypeName, string typeName, string expectedSetTypeName)
        {
            var class1 =
                $@"
Public Function Foo(baz As Long) As {typeName}
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {accessedTypeName}
    Dim baz As Variant
    Set baz = cls(42) 
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 22);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object", null)]
        [TestCase("Class1", "Object", null)]
        [TestCase("Variant", "Variant", null)]
        [TestCase("Class1", "Variant", null)]
        [TestCase("Object", "Class1", null)]
        [TestCase("Variant", "Class1", null)]
        [TestCase("Class1", "Class1", "TestProject.Class1")]
        [TestCase("Class1", "Long", null)]
        [TestCase("Object", "Long", null)]
        [TestCase("Variant", "Long", null)]
        public void DefaultMemberIndexExpression_SetTypeDeclarationTests(string accessedTypeName, string typeName, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                $@"
Public Function Foo(baz As Long) As {typeName}
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {accessedTypeName}
    Dim baz As Variant
    Set baz = cls(42) 
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 22);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("Variant", "Variant")]
        [TestCase("Long", SetTypeResolver.NotAnObject)]
        public void FunctionIndexExpression_SetTypeNameTests(string typeName, string expectedSetTypeName)
        {
            var class1 =
                @"
Public Function Foo(baz As Long) As Variant
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz As Variant
    Set baz = Foo(42) 
End Sub

Private Function Foo(baz As Long) As  {typeName}
End Function
";

            var expressionSelection = new Selection(4, 15, 4, 22);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [Category("ExpressionResolver")]
        [TestCase("Object", null)]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("Variant", null)]
        [TestCase("Long", null)]
        public void FunctionIndexExpression_SetTypeDeclarationTests(string typeName, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                @"
Public Function Foo(baz As Long) As Variant
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz As Variant
    Set baz = Foo(42) 
End Sub

Private Function Foo(baz As Long) As  {typeName}
End Function
";

            var expressionSelection = new Selection(4, 15, 4, 22);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("Variant", "Variant")]
        [TestCase("Long", SetTypeResolver.NotAnObject)]
        public void ArrayIndexExpression_SetTypeNameTests(string typeName, string expectedSetTypeName)
        {
            var class1 =
                @"
Public Function Foo(baz As Long) As Variant
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim arr(0 To 123) As {typeName}
    Dim baz As Variant
    Set baz = arr(42) 
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 22);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [Category("ExpressionResolver")]
        [TestCase("Object", null)]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("Variant", null)]
        [TestCase("Long", null)]
        public void ArrayIndexExpression_SetTypeDeclarationTests(string typeName, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                @"
Public Function Foo(baz As Long) As Variant
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim arr(0 To 123) As {typeName}
    Dim baz As Variant
    Set baz = arr(42) 
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 22);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object", null)]
        [TestCase("Class1", "Object", "Object")]
        [TestCase("Variant", "Variant", null)]
        [TestCase("Class1", "Variant", "Variant")]
        [TestCase("Object", "Class1", null)]
        [TestCase("Variant", "Class1", null)]
        [TestCase("Class1", "Class1", "TestProject.Class1")]
        [TestCase("Class1", "Long", SetTypeResolver.NotAnObject)]
        [TestCase("Object", "Long", null)]
        [TestCase("Variant", "Long", null)]
        public void DefaultMemberWhitespaceIndexExpression_SetTypeNameTests(string accessedTypeName, string typeName, string expectedSetTypeName)
        {
            var class1 =
                $@"
Public Function Foo(baz As Long) As {typeName}
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {accessedTypeName}
    Dim baz As Variant
    Set baz = cls _
 (42) 
End Sub
";

            var expressionSelection = new Selection(5, 15, 6, 6);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object", null)]
        [TestCase("Class1", "Object", null)]
        [TestCase("Variant", "Variant", null)]
        [TestCase("Class1", "Variant", null)]
        [TestCase("Object", "Class1", null)]
        [TestCase("Variant", "Class1", null)]
        [TestCase("Class1", "Class1", "TestProject.Class1")]
        [TestCase("Class1", "Long", null)]
        [TestCase("Object", "Long", null)]
        [TestCase("Variant", "Long", null)]
        public void DefaultMemberWhitespaceIndexExpression_SetTypeDeclarationTests(string accessedTypeName, string typeName, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                $@"
Public Function Foo(baz As Long) As {typeName}
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim cls As {accessedTypeName}
    Dim baz As Variant
    Set baz = cls _
 (42) 
End Sub
";

            var expressionSelection = new Selection(5, 15, 6, 6);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", "Object")]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("Variant", "Variant")]
        [TestCase("Long", SetTypeResolver.NotAnObject)]
        public void FunctionWhitespaceIndexExpression_SetTypeNameTests(string typeName, string expectedSetTypeName)
        {
            var class1 =
                @"
Public Function Foo(baz As Long) As Variant
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz As Variant
    Set baz = Foo _
 (42)
End Sub

Private Function Foo(baz As Long) As  {typeName}
End Function
";

            var expressionSelection = new Selection(4, 15, 5, 6);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        [TestCase("Object", null)]
        [TestCase("Class1", "TestProject.Class1")]
        [TestCase("Variant", null)]
        [TestCase("Long", null)]
        public void FunctionWhitespaceIndexExpression_SetTypeDeclarationTests(string typeName, string expectedNameOfSetTypeDeclaration)
        {
            var class1 =
                @"
Public Function Foo(baz As Long) As Variant
Attribute Foo.VB_UserMemId = 0
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz As Variant
    Set baz = Foo(42) 
End Sub

Private Function Foo(baz As Long) As  {typeName}
End Function
";

            var expressionSelection = new Selection(4, 15, 4, 22);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        public void IndexedDefaultMemberCallOnFunctionReturnValueHasTypeOfDefaultMember_SetTypeName()
        {
            var class1 =
                @"
Public Function Foo(baz As Long) As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";
            var class2 =
                @"
Private Function Foo() As Class1
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz As Variant
    Dim foo As Class1
    Set baz = foo.Foo(42) 
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 26);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Class2", ComponentType.ClassModule, class2)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var expectedSetTypeName = "TestProject.Class2";
            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        public void IndexedDefaultMemberCallOnFunctionReturnValueHasTypeOfDefaultMember_SetTypeDeclaration()
        {
            var class1 =
                @"
Public Function Foo(baz As Long) As Class2
Attribute Foo.VB_UserMemId = 0
End Function
";
            var class2 =
                @"
Private Function Foo() As Class1
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz As Variant
    Dim foo As Class1
    Set baz = foo.Foo(42) 
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 26);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Class2", ComponentType.ClassModule, class2)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);

            var expectedNameOfSetTypeDeclaration = "TestProject.Class2";
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        [Test]
        [Category("ExpressionResolver")]
        public void ArrayAccessOnFunctionReturnValueHasTypeOfDefaultMember_SetTypeName()
        {
            var class1 =
                @"
Public Function Foo() As Class2()
Attribute Foo.VB_UserMemId = 0
End Function
";
            var class2 =
                @"
Private Function Foo() As Class1
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz As Variant
    Dim foo As Class1
    Set baz = foo.Foo(42) 
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 26);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Class2", ComponentType.ClassModule, class2)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var expectedSetTypeName = "TestProject.Class2";
            var actualSetTypeName = ExpressionTypeName(vbe, "Module1", expressionSelection);

            Assert.AreEqual(expectedSetTypeName, actualSetTypeName);
        }

        [Test]
        [Category("ExpressionResolver")]
        public void ArrayAccessOnFunctionReturnValueHasTypeOfDefaultMember_SetTypeDeclaration()
        {
            var class1 =
                @"
Public Function Foo() As Class2()
Attribute Foo.VB_UserMemId = 0
End Function
";
            var class2 =
                @"
Private Function Foo() As Class1
End Function
";

            var module1 =
                $@"
Private Sub Bar()
    Dim baz As Variant
    Dim foo As Class1
    Set baz = foo.Foo(42) 
End Sub
";

            var expressionSelection = new Selection(5, 15, 5, 26);

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Class2", ComponentType.ClassModule, class2)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeDeclaration = ExpressionTypeDeclaration(vbe, "Module1", expressionSelection);

            var expectedNameOfSetTypeDeclaration = "TestProject.Class2";
            var actualNameOfSetTypeDeclaration = setTypeDeclaration?.QualifiedModuleName.ToString();

            Assert.AreEqual(expectedNameOfSetTypeDeclaration, actualNameOfSetTypeDeclaration);
        }

        private Declaration ExpressionTypeDeclaration(IVBE vbe, string componentName, Selection selection)
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var resolver = ExpressionResolverUnderTest(state);
                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName.Equals(componentName));
                var expression = TestExpression(state, module, selection);
                return resolver.SetTypeDeclaration(expression, module);
            }
        }

        private string ExpressionTypeName(IVBE vbe, string componentName, Selection selection)
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var resolver = ExpressionResolverUnderTest(state);
                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName.Equals(componentName));
                var expression = TestExpression(state, module, selection);
                return resolver.SetTypeName(expression, module);
            }
        }

        private VBAParser.ExpressionContext TestExpression(IParseTreeProvider parseTreeProvider, QualifiedModuleName module, Selection selection)
        {
            if (!(parseTreeProvider.GetParseTree(module, CodeKind.CodePaneCode) is ParserRuleContext context))
            {
                return null;
            }

            if (!context.GetSelection().Contains(selection))
            {
                return null;
            }

            return TestExpression(context, selection);
        }

        private VBAParser.ExpressionContext TestExpression(ParserRuleContext context, Selection selection)
        {
            if (context == null)
            {
                return null;
            }

            var containingChild = context.children
                .OfType<ParserRuleContext>()
                .FirstOrDefault(childContext => childContext.GetSelection().Contains(selection));

            var containedTestExpression = containingChild != null
                ? TestExpression(containingChild, selection)
                : null;

            if (containedTestExpression != null)
            {
                return containedTestExpression;
            }

            if (context is VBAParser.ExpressionContext expression)
            {
                return expression;
            }

            return null;
        }

        private static ISetTypeResolver ExpressionResolverUnderTest(IDeclarationFinderProvider declarationFinderProvider)
        {
            return new SetTypeResolver(declarationFinderProvider);
        }
    }
}
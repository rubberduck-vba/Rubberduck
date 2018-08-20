using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using RubberduckTests.Mocks;

namespace RubberduckTests.Grammar
{
    [TestFixture]
    public class MembersAllowingAttributesTests
    {
        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ModuleVariablesInADeclarationListAreRecognziedToAllowingAttributes()
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
    Dim wooBaz As Integer
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
                var expectedAttributeScope = ("bar", DeclarationType.Variable);
                var otherExpectedAttributeScope = ("fooBazz", DeclarationType.Variable);

                var membersAllowingAttributes = state.GetMembersAllowingAttributes(module);
                Assert.IsTrue(membersAllowingAttributes.ContainsKey(expectedAttributeScope));
                Assert.IsTrue(membersAllowingAttributes.ContainsKey(otherExpectedAttributeScope));
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void LocalVariablesInADeclarationListAreNotRecognziedToAllowingAttributes()
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
    Dim wooBaz, hrm As Integer
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
                var notExpectedAttributeScope = ("wooBaz", DeclarationType.Variable);
                var otherNotExpectedAttributeScope = ("hrm", DeclarationType.Variable);

                var membersAllowingAttributes = state.GetMembersAllowingAttributes(module);
                Assert.IsFalse(membersAllowingAttributes.ContainsKey(notExpectedAttributeScope));
                Assert.IsFalse(membersAllowingAttributes.ContainsKey(otherNotExpectedAttributeScope));
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void SingleModuleVariablesAreRecognziedToAllowingAttributes()
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
    Dim wooBaz As Integer
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
                var expectedAttributeScope = ("baz", DeclarationType.Variable);

                var membersAllowingAttributes = state.GetMembersAllowingAttributes(module);
                Assert.IsTrue(membersAllowingAttributes.ContainsKey(expectedAttributeScope));
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void SingleLocalVariablesAreNotRecognziedToAllowingAttributes()
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
    Dim wooBaz As Integer
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
                var notExpectedAttributeScope = ("wooBaz", DeclarationType.Variable);

                var membersAllowingAttributes = state.GetMembersAllowingAttributes(module);
                Assert.IsFalse(membersAllowingAttributes.ContainsKey(notExpectedAttributeScope));
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void ProceduresAreRecognziedToAllowingAttributes()
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
    Dim wooBaz As Integer
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
                var expectedAttributeScope = ("Foo", DeclarationType.Procedure);

                var membersAllowingAttributes = state.GetMembersAllowingAttributes(module);
                Assert.IsTrue(membersAllowingAttributes.ContainsKey(expectedAttributeScope));
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void FunctionsAreRecognziedToAllowingAttributes()
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
    Dim wooBaz As Integer
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
                var expectedAttributeScope = ("Woo", DeclarationType.Function);

                var membersAllowingAttributes = state.GetMembersAllowingAttributes(module);
                Assert.IsTrue(membersAllowingAttributes.ContainsKey(expectedAttributeScope));
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertyGetsAreRecognziedToAllowingAttributes()
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
    Dim wooBaz As Integer
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
                var expectedAttributeScope = ("FooBar", DeclarationType.PropertyGet);

                var membersAllowingAttributes = state.GetMembersAllowingAttributes(module);
                Assert.IsTrue(membersAllowingAttributes.ContainsKey(expectedAttributeScope));
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertyLetsAreRecognziedToAllowingAttributes()
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
    Dim wooBaz As Integer
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
                var expectedAttributeScope = ("FooBar", DeclarationType.PropertyLet);

                var membersAllowingAttributes = state.GetMembersAllowingAttributes(module);
                Assert.IsTrue(membersAllowingAttributes.ContainsKey(expectedAttributeScope));
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [Test]
        public void PropertySetsAreRecognziedToAllowingAttributes()
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
    Dim wooBaz As Integer
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
                var expectedAttributeScope = ("FooBaz", DeclarationType.PropertySet);

                var membersAllowingAttributes = state.GetMembersAllowingAttributes(module);
                Assert.IsTrue(membersAllowingAttributes.ContainsKey(expectedAttributeScope));
            }
        }
    }
}
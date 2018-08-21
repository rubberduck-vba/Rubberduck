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
        [TestCase("bar", DeclarationType.Variable)]
        [TestCase("fooBazz", DeclarationType.Variable)]
        [TestCase("baz", DeclarationType.Variable)]
        [TestCase("Foo", DeclarationType.Procedure)]
        [TestCase("Woo", DeclarationType.Function)]
        [TestCase("FooBar", DeclarationType.PropertyGet)]
        [TestCase("FooBar", DeclarationType.PropertyLet)]
        [TestCase("FooBaz", DeclarationType.PropertySet)]
        public void MembersAllowingAttributesAreRecognized(string memberName, DeclarationType declarationType)
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
    Dim wooBaz As Integer
    Dim barFoo, bazFoo As String
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
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScope = (memberName, declarationType);

                var membersAllowingAttributes = state.GetMembersAllowingAttributes(module);
                Assert.IsTrue(membersAllowingAttributes.ContainsKey(expectedAttributeScope));
            }
        }

        [Category("Grammar")]
        [Category("Attributes")]
        [TestCase("wooBaz", DeclarationType.Variable)]
        [TestCase("barFoo", DeclarationType.Variable)]
        [TestCase("bazFoo", DeclarationType.Variable)]
        public void LocalVariableDeclarationsAreNotRecognizedToAllowAttributes(string memberName, DeclarationType declarationType)
        {
            const string inputCode =
                @"

Public bar As Long, fooBazz As Integer

Public baz As String

Public Sub Foo(ByRef arg1 As String)
    arg1 = ""test""
    Dim wooBaz As Integer
    Dim barFoo, bazFoo As String
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
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            var module = component.QualifiedModuleName;
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var expectedAttributeScope = (memberName, declarationType);

                var membersAllowingAttributes = state.GetMembersAllowingAttributes(module);
                Assert.IsFalse(membersAllowingAttributes.ContainsKey(expectedAttributeScope));
            }
        }
    }
}
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using System.Linq;
using System.Threading;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Binding
{

    [TestFixture]
    public class IndexDefaultBindingTests
    {
        private const string BINDING_TARGET_LEXPRESSION = "BindingTarget";
        private const string BINDING_TARGET_UNRESTRICTEDNAME = "UnrestrictedName";
        private const string TEST_CLASS_NAME = "TestClass";
        private const string REFERENCED_PROJECT_FILEPATH = @"C:\Temp\ReferencedProjectA";

        [Category("Binding")]
        [Test]
        public void IndexedDefaultMember()
        {
            string callerModule = @"
Public Property Get Ok() As Klasse2
End Property

Public Sub Test()
    Call Ok(1)
End Sub
";

            string defaultMemberClass = @"
Public Property Get Test2(a As Integer)
Attribute Test2.VB_UserMemId = 0
    Test2 = 2
End Property
";

            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("Any Project", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent("AnyModule1", ComponentType.StandardModule, callerModule);
            enclosingProjectBuilder.AddComponent("Klasse2", ComponentType.ClassModule, defaultMemberClass);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {
                var declaration =
                    state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.PropertyGet &&
                                                          d.IdentifierName == "Test2");

                //One for the assignment in the property itself and one for the default member access.
                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void IndexedDefaultMemberMemberAccess()
        {
            string callerModule = @"
Public Property Get Ok() As Klasse2
End Property

Public Sub Test()
    Dim bar As Long
    bar = Ok(1).Test3
End Sub
";

            string defaultMemberClass = @"
Public Property Get Test2(a As Integer) As Klasse3
Attribute Test2.VB_UserMemId = 0
    Set Test2 = New Klasse3
End Property
";
            string defaultMemberTargetClass = @"
Public Property Get Test3() As Long
Attribute Test3.VB_UserMemId = 0
    Test3 = 3
End Property
";

            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("Any Project", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent("AnyModule1", ComponentType.StandardModule, callerModule);
            enclosingProjectBuilder.AddComponent("Klasse2", ComponentType.ClassModule, defaultMemberClass);
            enclosingProjectBuilder.AddComponent("Klasse3", ComponentType.ClassModule, defaultMemberTargetClass);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {
                var declaration =
                    state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.PropertyGet &&
                                                          d.IdentifierName == "Test3");

                //One for the assignment in the property itself and one for the access via the default member.
                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void RecursiveDefaultMember()
        {
            string callerModule = @"
Public Property Get Ok() As Klasse1
End Property

Public Sub Test()
    Call Ok(1)
End Sub
";

            string middleman = @"
Public Property Get Test1() As Klasse2
Attribute Test1.VB_UserMemId = 0
End Property
";

            string defaultMemberClass = @"
Public Property Get Test2(a As Integer)
    Attribute Test2.VB_UserMemId = 0
    Test2 = 2
End Property
";

            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("Any Project", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent("AnyModule1", ComponentType.StandardModule, callerModule);
            enclosingProjectBuilder.AddComponent("Klasse1", ComponentType.ClassModule, middleman);
            enclosingProjectBuilder.AddComponent("Klasse2", ComponentType.ClassModule, defaultMemberClass);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {
                var declaration =
                    state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.PropertyGet &&
                                                          d.IdentifierName == "Test2");

                //One for the assignment in the property itself and one for the default member access.
                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void RecursiveDefaultMemberMemberAccess()
        {
            string callerModule = @"
Public Property Get Ok() As Klasse1
End Property

Private Function Foo(a As Integer)
    Foo = 5
End Function

Public Sub Test()
    Dim bar As Long
    bar = Foo(Ok(1).Test3)
End Sub
";

            string middleman = @"
Public Property Get Test1() As Klasse2
Attribute Test1.VB_UserMemId = 0
End Property
";

            string defaultMemberClass = @"
Public Property Get Test2(a As Integer) As Klasse3
Attribute Test2.VB_UserMemId = 0
    Set Test2 = New Klasse3
End Property
";
            string defaultMemberTargetClass = @"
Public Property Get Test3() As Long
Attribute Test3.VB_UserMemId = 0
    Test3 = 3
End Property
";

            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("Any Project", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent("AnyModule1", ComponentType.StandardModule, callerModule);
            enclosingProjectBuilder.AddComponent("Klasse1", ComponentType.ClassModule, middleman);
            enclosingProjectBuilder.AddComponent("Klasse2", ComponentType.ClassModule, defaultMemberClass);
            enclosingProjectBuilder.AddComponent("Klasse3", ComponentType.ClassModule, defaultMemberTargetClass);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {
                var declaration =
                    state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.PropertyGet &&
                                                          d.IdentifierName == "Test3");

                //One for the assignment in the property itself and one for the access via the default member.
                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void RecursiveIndexedDefaultMember()
        {
            string callerModule = @"
Public Property Get Ok() As Klasse1
End Property

Public Sub Test()
    Call Ok(1)(2)
End Sub
";

            string middleman = @"
Public Property Get Test1(bar As Long) As Klasse2
Attribute Test1.VB_UserMemId = 0
End Property
";

            string defaultMemberClass = @"
Public Property Get Test2(a As Integer)
    Attribute Test2.VB_UserMemId = 0
    Test2 = 2
End Property
";

            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("Any Project", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent("AnyModule1", ComponentType.StandardModule, callerModule);
            enclosingProjectBuilder.AddComponent("Klasse1", ComponentType.ClassModule, middleman);
            enclosingProjectBuilder.AddComponent("Klasse2", ComponentType.ClassModule, defaultMemberClass);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {
                var declaration =
                    state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.PropertyGet &&
                                                          d.IdentifierName == "Test2");

                //One for the assignment in the property itself and one for the default member access.
                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void RecursiveIndexedDefaultMemberMemberAccess()
        {
            string callerModule = @"
Public Property Get Ok() As Klasse1
End Property

Private Sub DoIt(a As Long)
End Sub

Public Sub Test()
    DoIt Ok(1)(2).Test3
End Sub
";

            string middleman = @"
Public Property Get Test1(bar As Long) As Klasse2
Attribute Test1.VB_UserMemId = 0
End Property
";

            string defaultMemberClass = @"
Public Property Get Test2(a As Integer) As Klasse3
Attribute Test2.VB_UserMemId = 0
    Set Test2 = New Klasse3
End Property
";
            string defaultMemberTargetClass = @"
Public Property Get Test3() As Long
Attribute Test3.VB_UserMemId = 0
    Test3 = 3
End Property
";

            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("Any Project", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent("AnyModule1", ComponentType.StandardModule, callerModule);
            enclosingProjectBuilder.AddComponent("Klasse1", ComponentType.ClassModule, middleman);
            enclosingProjectBuilder.AddComponent("Klasse2", ComponentType.ClassModule, defaultMemberClass);
            enclosingProjectBuilder.AddComponent("Klasse3", ComponentType.ClassModule, defaultMemberTargetClass);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();
            using (var state = Parse(vbe))
            {
                var declaration =
                    state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.PropertyGet &&
                                                          d.IdentifierName == "Test3");

                //One for the assignment in the property itself and one for the access via the default member.
                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Binding")]
        [Test]
        public void NormalPropertyFunctionSubroutine()
        {
            string callerModule = @"
Public Sub Test()
    Call Test1(1)
End Sub
";

            string property = @"
Public Property Get Test1(a As Integer) As Long
End Property
";

            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder("Any Project", ProjectProtection.Unprotected);
            enclosingProjectBuilder.AddComponent("AnyModule1", ComponentType.StandardModule, callerModule);
            enclosingProjectBuilder.AddComponent("AnyClass", ComponentType.StandardModule, property);
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            var vbe = builder.Build();

            using (var state = Parse(vbe))
            {
                var declaration =
                    state.AllUserDeclarations.Single(d => d.DeclarationType == DeclarationType.PropertyGet &&
                                                          d.IdentifierName == "Test1");

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        private static RubberduckParserState Parse(Mock<IVBE> vbe)
        {
            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }
            var state = parser.State;
            return state;
        }
    }
}

using System;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.Annotations;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace RubberduckTests.Grammar
{
    [TestFixture]
    public class ResolverTests
    {
        private RubberduckParserState Resolve(string code, bool loadStdLib = false, ComponentType moduleType = ComponentType.StandardModule)
        {
            IVBComponent component;
            var vbe = MockVbeBuilder.BuildFromSingleModule(code, moduleType, out component, Selection.Empty, loadStdLib);
            var parser = MockParser.Create(vbe.Object);
            var state = parser.State;
            parser.Parse(new CancellationTokenSource());

            if (state.Status == ParserState.ResolverError)
            {
                Assert.Fail("Parser state should be 'Ready', but returns '{0}'.", state.Status);
            }
            if (state.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", state.Status);
            }

            return state;
        }

        private RubberduckParserState Resolve(params string[] classes)
        {
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            for (var i = 0; i < classes.Length; i++)
            {
                projectBuilder.AddComponent("Class" + (i + 1), ComponentType.ClassModule, classes[i]);
            }

            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            var parser = MockParser.Create(vbe.Object);
            var state = parser.State;
            parser.Parse(new CancellationTokenSource());

            if (state.Status == ParserState.ResolverError)
            {
                Assert.Fail("Parser state should be 'Ready', but returns '{0}'.", state.Status);
            }
            if (state.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", state.Status);
            }

            return state;
        }

        private RubberduckParserState Resolve(params Tuple<string, ComponentType>[] components)
        {
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected);
            for (var i = 0; i < components.Length; i++)
            {
                projectBuilder.AddComponent("Component" + (i + 1), components[i].Item2, components[i].Item1);
            }

            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();
            var parser = MockParser.Create(vbe.Object);
            var state = parser.State;
            parser.Parse(new CancellationTokenSource());

            if (state.Status == ParserState.ResolverError)
            {
                Assert.Fail("Parser state should be 'Ready', but returns '{0}'.", state.Status);
            }
            if (state.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", state.Status);
            }

            return state;
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void FunctionReturnValueAssignment_IsReferenceToFunctionDeclaration()
        {
            var code = @"
Public Function Foo() As String
    Foo = 42
End Function
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Function && item.IdentifierName == "Foo");

                Assert.AreEqual(1, declaration.References.Count(item => item.IsAssignment));
            }
        }

        [Category("Resolver")]
        [Test]
        public void JaggedArrayReference_DoesNotBlowUp()
        {
            // see https://github.com/rubberduck-vba/Rubberduck/issues/3098
            var code = @"Option Explicit

Public Sub Test()
    Dim varTemp() As Variant
    
    ReDim varTemp(0)
    
    varTemp(0) = Array(0)
    varTemp(0)(0) = Array(0)
    
    Debug.Print varTemp(0)(0)
End Sub
";

            using (var state = Resolve(code))
            {
                var declaration = state.AllUserDeclarations.Single(item => item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "varTemp");
                Assert.IsTrue(declaration.IsArray);
            }
        }

        [Category("Resolver")]
        [Test]
        public void OptionalParameterDefaultConstValue_IsReferenceToDeclaredConst()
        {
            var code = @"
Public Const Foo As Long = 42
Public Sub DoSomething(Optional ByVal bar As Long = Foo)
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Constant && item.IdentifierName == "Foo");

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void TypeOfIsExpression_BooleanExpressionIsReferenceToLocalVariable()
        {
            var code_class1 = @"
Public Function Foo() As String
    Dim a As Object
    anything = TypeOf a Is Class2
End Function
";
            // We only use the second class as as target of the type expression, its contents don't matter.
            var code_class2 = string.Empty;

            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "a");

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void TypeOfIsExpression_TypeExpressionIsReferenceToClass()
        {
            var code_class1 = @"
Public Function Foo() As String
    Dim a As Object
    anything = TypeOf a Is Class2
End Function
";
            // We only use the second class as as target of the type expression, its contents don't matter.
            var code_class2 = string.Empty;

            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.ClassModule && item.IdentifierName == "Class2");

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void FunctionCall_IsReferenceToFunctionDeclaration()
        {
            var code = @"
Public Sub DoSomething()
    Foo
End Sub

Private Function Foo() As String
    Foo = 42
End Function
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Function && item.IdentifierName == "Foo");

                var reference = declaration.References.SingleOrDefault(item => !item.IsAssignment);
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void FunctionCallWithParensOnNextContinuedLine_IsReferenceToFunctionDeclaration()
        {
            var code = @"
Public Sub DoSomething()
    Bar Foo _
  ()
End Sub

Public Sub Bar()
End Sub

Private Function Foo() As String
    Foo = 42
End Function
";
            using (var state = Resolve(code))
            {

                var declaration = state.DeclarationFinder
                    .UserDeclarations(DeclarationType.Function)
                    .Single(item => item.IdentifierName == "Foo");

                var reference = declaration.References.SingleOrDefault(item => !item.IsAssignment);
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void LocalVariableCall_IsReferenceToVariableDeclaration()
        {
            var code = @"
Public Sub DoSomething()
    Dim foo As Integer
    a = foo
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

                var reference = declaration.References.SingleOrDefault(item => !item.IsAssignment);
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void LocalVariableForeignNameCall_IsReferenceToVariableDeclaration()
        {
            var code = @"
Public Sub DoSomething()
    Dim foo As Integer
    a = [foo]
End Sub
";
            using (var state = Resolve(code, true))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

                var reference = declaration.References.SingleOrDefault(item => !item.IsAssignment);
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void LocalVariableAssignment_IsReferenceToVariableDeclaration()
        {
            var code = @"
Public Sub DoSomething()
    Dim foo As Integer
    foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

                var reference = declaration.References.SingleOrDefault(item => item.IsAssignment);
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void SingleLineIfStatementLabel_IsReferenceToLabel_NumberLabelHasColon()
        {
            var code = @"
Public Sub DoSomething()
    If True Then 5
5:
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.LineLabel && item.IdentifierName == "5");

                var reference = declaration.References.SingleOrDefault();
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void SingleLineIfStatementLabel_IsReferenceToLabel_NumberLabelNoColon()
        {
            var code = @"
Public Sub DoSomething()
    Dim fizz As Integer
    fizz = 5
    If fizz = 5 Then Exit Sub

    If True Then 5
5
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.LineLabel && item.IdentifierName == "5");

                var reference = declaration.References.SingleOrDefault();
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void SingleLineIfStatementLabel_IsReferenceToLabel_IdentifierLabel()
        {
            var code = @"
Public Sub DoSomething()
    If True Then GoTo foo
foo:
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.LineLabel && item.IdentifierName == "foo");

                var reference = declaration.References.SingleOrDefault();
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ProjectUdtSameNameFirstProjectThenUdt_FirstReferenceIsToProject()
        {
            var code = string.Format(@"
Private Type {0}
    anything As String
End Type

Public Sub DoSomething()
    Dim a As {0}.{1}.{0}
End Sub
", MockVbeBuilder.TestProjectName, MockVbeBuilder.TestModuleName);
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Project && item.IdentifierName == MockVbeBuilder.TestProjectName);

                var reference = declaration.References.SingleOrDefault();
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ProjectUdtSameNameUdtOnly_IsReferenceToUdt()
        {
            var code = string.Format(@"
Private Type {0}
    anything As String
End Type

Public Sub DoSomething()
    Dim a As {1}.{0}
End Sub
", MockVbeBuilder.TestProjectName, MockVbeBuilder.TestModuleName);
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedType && item.IdentifierName == MockVbeBuilder.TestProjectName);

                var reference = declaration.References.SingleOrDefault();
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void EncapsulatedVariableAssignment_DoesNotResolve()
        {
            var code_class1 = @"
Public Sub DoSomething()
    foo = 42
End Sub
";
            var code_class2 = @"
Option Explicit
Public foo As Integer
";
            var class1 = Tuple.Create(code_class1, ComponentType.ClassModule);
            var class2 = Tuple.Create(code_class2, ComponentType.ClassModule);

            using (var state = Resolve(class1, class2))
            {

                var declaration = state.AllUserDeclarations.Single(item => item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo" && !item.IsUndeclared);

                var reference = declaration.References.SingleOrDefault(item => item.IsAssignment);
                Assert.IsNull(reference);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void PublicVariableCall_IsReferenceToVariableDeclaration()
        {
            var code_class1 = @"
Public Sub DoSomething()
    a = foo
End Sub
";
            var code_class2 = @"
Option Explicit
Public foo As Integer
";
            using (var state = Resolve(
                Tuple.Create(code_class1, ComponentType.ClassModule),
                Tuple.Create(code_class2, ComponentType.StandardModule)))
            {
                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

                var reference = declaration.References.SingleOrDefault(item => !item.IsAssignment);
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void PublicVariableAssignment_IsReferenceToVariableDeclaration()
        {
            var code_class1 = @"
Public Sub DoSomething()
    foo = 42
End Sub
";
            var code_module1 = @"
Option Explicit
Public foo As Integer
";
            var class1 = Tuple.Create(code_class1, ComponentType.ClassModule);
            var module1 = Tuple.Create(code_module1, ComponentType.StandardModule);

            using (var state = Resolve(class1, module1))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

                var reference = declaration.References.SingleOrDefault(item => item.IsAssignment);
                Assert.IsNotNull(reference);
                Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void UserDefinedTypeVariableAsTypeClause_IsReferenceToUserDefinedTypeDeclaration()
        {
            var code = @"
Private Type TFoo
    Bar As Integer
End Type
Private this As TFoo
";
            using (var state = Resolve(code, false, ComponentType.ClassModule))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedType && item.IdentifierName == "TFoo");

                Assert.IsNotNull(declaration.References.SingleOrDefault());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ObjectVariableAsTypeClause_IsReferenceToClassModuleDeclaration()
        {
            var code_class1 = @"
Public Sub DoSomething()
    Dim foo As Class2
End Sub
";
            var code_class2 = @"
Option Explicit
";

            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.ClassModule && item.IdentifierName == "Class2");

                Assert.IsNotNull(declaration.References.SingleOrDefault());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ParameterCall_IsReferenceToParameterDeclaration()
        {
            var code = @"
Public Sub DoSomething(ByVal foo As Integer)
    a = foo
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Parameter && item.IdentifierName == "foo");

                Assert.IsNotNull(declaration.References.SingleOrDefault());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ParameterAssignment_IsAssignmentReferenceToParameterDeclaration()
        {
            var code = @"
Public Sub DoSomething(ByRef foo As Integer)
    foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Parameter && item.IdentifierName == "foo");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item => item.IsAssignment));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void NamedParameterCall_IsReferenceToParameterDeclaration()
        {
            var code = @"
Public Sub DoSomething()
    DoSomethingElse foo:=42
End Sub

Private Sub DoSomethingElse(ByVal foo As Integer)
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Parameter && item.IdentifierName == "foo");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.IdentifierName == "DoSomething"));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void UserDefinedTypeMemberCall_IsReferenceToUserDefinedTypeMemberDeclaration()
        {
            var code = @"
Private Type TFoo
    Bar As Integer
End Type
Private this As TFoo

Public Property Get Bar() As Integer
    Bar = this.Bar
End Property
";
            using (var state = Resolve(code, false, ComponentType.ClassModule))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedTypeMember && item.IdentifierName == "Bar");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.PropertyGet
                    && item.ParentScoping.IdentifierName == "Bar"));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void UserDefinedTypeVariableCall_IsReferenceToVariableDeclaration()
        {
            var code = @"
Private Type TFoo
    Bar As Integer
End Type
Private this As TFoo

Public Property Get Bar() As Integer
    Bar = this.Bar
End Property
";
            using (var state = Resolve(code, false, ComponentType.ClassModule))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "this");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.PropertyGet
                    && item.ParentScoping.IdentifierName == "Bar"));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void WithVariableMemberCall_IsReferenceToMemberDeclaration()
        {
            var code_class1 = @"
Public Property Get Foo() As Integer
    Foo = 42
End Property
";
            var code_class2 = @"
Public Sub DoSomething()
    With New Class1
        a = .Foo
    End With
End Sub
";
            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.PropertyGet && item.IdentifierName == "Foo");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.Procedure
                    && item.ParentScoping.IdentifierName == "DoSomething"));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void NestedWithVariableMemberCall_IsReferenceToMemberDeclaration()
        {
            var code_class1 = @"
Public Property Get Foo() As Class2
    Foo = New Class2
End Property
";
            var code_class2 = @"
Public Property Get Bar() As Integer
    Bar = 42
End Property
";
            var code_class3 = @"
Public Sub DoSomething()
    With New Class1
        With .Foo
            a = .Bar
        End With
    End With
End Sub
";

            using (var state = Resolve(code_class1, code_class2, code_class3))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.PropertyGet && item.IdentifierName == "Bar");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.Procedure
                    && item.ParentScoping.IdentifierName == "DoSomething"));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ResolvesLocalVariableToSmallestScopeIdentifier()
        {
            var code = @"
Private foo As Integer

Private Sub DoSomething()
    Dim foo As Integer
    foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable
                    && item.ParentScopeDeclaration.IdentifierName == "DoSomething"
                    && item.IdentifierName == "foo");

                Assert.IsNotNull(declaration.References.SingleOrDefault());

                var fieldDeclaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable
                    && item.ParentScopeDeclaration.DeclarationType == DeclarationType.ProceduralModule
                    && item.IdentifierName == "foo");

                Assert.IsNull(fieldDeclaration.References.SingleOrDefault());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void Implements_IsReferenceToClassDeclaration()
        {
            var code_class1 = @"
Public Sub DoSomething()
End Sub
";
            var code_class2 = @"
Implements Class1

Private Sub Class1_DoSomething()
End Sub
";
            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.ClassModule && item.IdentifierName == "Class1");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.IdentifierName == "Class2"));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void NestedMemberCall_IsReferenceToMember()
        {
            var code_class1 = @"
Public Property Get Foo() As Class2
    Foo = New Class2
End Property
";
            var code_class2 = @"
Public Property Get Bar() As Integer
    Bar = 42
End Property
";
            var code_class3 = @"
Public Sub DoSomething(ByVal a As Class1)
    a = a.Foo.Bar
End Sub
";
            using (var state = Resolve(code_class1, code_class2, code_class3))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.PropertyGet && item.IdentifierName == "Bar");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.Procedure
                    && item.ParentScoping.IdentifierName == "DoSomething"));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void MemberCallParent_IsReferenceToParent()
        {
            var code_class1 = @"
Public Property Get Foo() As Integer
    Foo = 42
End Property
";
            var code_class2 = @"
Public Sub DoSomething(ByVal a As Class1)
    b = a.Foo
End Sub
";
            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Parameter && item.IdentifierName == "a");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.Procedure
                    && item.ParentScoping.IdentifierName == "DoSomething"));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ForLoop_IsAssignmentReferenceToIteratorDeclaration()
        {
            var code = @"
Public Sub DoSomething()
    Dim i As Integer
    For i = 0 To 9
    Next
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "i");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.Procedure
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && item.IsAssignment));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ForLoop_AddsReferenceEvenIfAssignmentResolutionFailure()
        {
            var code = @"
Public Sub DoSomething()
    Dim i As Integer
    For i = doesntExist To doesntExistEither
    Next
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "i");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.Procedure
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && item.IsAssignment));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ForEachLoop_IsReferenceToIteratorDeclaration()
        {
            var code = @"
Public Sub DoSomething(ByVal c As Collection)
    Dim i As Variant
    For Each i In c
    Next
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "i");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.Procedure
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && item.IsAssignment));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ForEachLoop_InClauseIsReferenceToIteratedDeclaration()
        {
            var code = @"
Public Sub DoSomething(ByVal c As Collection)
    Dim i As Variant
    For Each i In c
    Next
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Parameter && item.IdentifierName == "c");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.Procedure
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && !item.IsAssignment));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ArraySubscriptAccess_IsReferenceToArrayDeclaration()
        {
            var code = @"
Public Sub DoSomething(ParamArray values())
    Dim i As Integer
    For i = 0 To 9
        a = values(i)
    Next
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Parameter
                    && item.IdentifierName == "values"
                    && item.IsArray);

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.Procedure
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && !item.IsAssignment));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void SubscriptWrite_IsNotAssignmentReferenceToObjectDeclaration()
        {
            var code = @"
Public Sub DoSomething()
    Dim foo As Object
    Set foo = CreateObject(""Scripting.Dictionary"")
    foo(""key"") = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable
                    && item.IdentifierName == "foo");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.Procedure
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && !item.IsAssignment));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ArraySubscriptWrite_IsAssignmentReferenceToArrayDeclaration()
        {
            var code = @"
Public Sub DoSomething(ParamArray values())
    Dim i As Integer
    For i = LBound(values) To UBound(values)
        values(i) = 42
    Next
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Parameter
                    && item.IdentifierName == "values"
                    && item.IsArray);

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.DeclarationType == DeclarationType.Procedure
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && item.IsAssignment));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void PropertyGetCall_IsReferenceToPropertyGetDeclaration()
        {
            var code_class1 = @"
Private tValue As Integer

Public Property Get Foo() As Integer
    Foo = tValue
End Property

Public Property Let Foo(ByVal value As Integer)
    tValue = value
End Property
";
            var code_class2 = @"
Public Sub DoSomething()
    Dim bar As New Class1
    a = bar.Foo
End Sub
";

            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.PropertyGet
                    && item.IdentifierName == "Foo");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    !item.IsAssignment
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void PropertySetCall_IsReferenceToPropertySetDeclaration()
        {
            var code_class1 = @"
Private tValue As Object

Public Property Get Foo() As Object
    Set Foo = tValue
End Property

Public Property Set Foo(ByVal value As Object)
    Set tValue = value
End Property
";
            var code_class2 = @"
Public Sub DoSomething()
    Dim bar As New Class1
    Set bar.Foo = Nothing
End Sub
";

            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.PropertySet
                    && item.IdentifierName == "Foo");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.IsAssignment
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void PropertyLetCall_IsReferenceToPropertyLetDeclaration()
        {
            var code_class1 = @"
Private tValue As Integer

Public Property Get Foo() As Integer
    Foo = tValue
End Property

Public Property Let Foo(ByVal value As Integer)
    tValue = value
End Property
";
            var code_class2 = @"
Public Sub DoSomething()
    Dim bar As New Class1
    bar.Foo = 42
End Sub
";

            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.PropertyLet
                    && item.IdentifierName == "Foo");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.IsAssignment
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void EnumMemberCall_IsReferenceToEnumMemberDeclaration()
        {
            var code = @"
Option Explicit
Public Enum FooBarBaz
    Foo
    Bar
    Baz
End Enum

Public Sub DoSomething()
    Dim a As FooBarBaz
    a = Foo
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.EnumerationMember
                    && item.IdentifierName == "Foo"
                    && item.ParentDeclaration.IdentifierName == "FooBarBaz");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    !item.IsAssignment
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && item.ParentScoping.DeclarationType == DeclarationType.Procedure));

            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void QualifiedEnumMemberCall_IsReferenceToEnumMemberDeclaration()
        {
            var code = @"
Option Explicit
Public Enum FooBarBaz
    Foo
    Bar
    Baz
End Enum

Public Sub DoSomething()
    Dim a As FooBarBaz
    a = FooBarBaz.Foo
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.EnumerationMember
                    && item.IdentifierName == "Foo"
                    && item.ParentDeclaration.IdentifierName == "FooBarBaz");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    !item.IsAssignment
                    && item.ParentScoping.IdentifierName == "DoSomething"
                    && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void EnumParameterAsTypeName_ResolvesToEnumTypeDeclaration()
        {
            var code = @"

Option Explicit
Public Enum FooBarBaz
    Foo
    Bar
    Baz
End Enum

Public Sub DoSomething(ByVal a As FooBarBaz)
End Sub
";

            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Enumeration
                    && item.IdentifierName == "FooBarBaz");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.IdentifierName == "DoSomething"
                    && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void FunctionWithSameNameAsEnumReturnAssignment_DoesntResolveToEnum()
        {
            var code = @"

Option Explicit
Public Enum Foos
    Foo1
End Enum

Public Function Foos() As Foos
    Foos = Foo1
End Function
";

            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Enumeration
                    && item.IdentifierName == "Foos");

                Assert.IsTrue(declaration.References.All(item => item.Selection.StartLine != 9));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void UserDefinedTypeParameterAsTypeName_ResolvesToUserDefinedTypeDeclaration()
        {
            var code = @"
Option Explicit
Public Type TFoo
    Foo As Integer
End Type

Public Sub DoSomething(ByVal a As TFoo)
End Sub
";

            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedType
                    && item.IdentifierName == "TFoo");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                    item.ParentScoping.IdentifierName == "DoSomething"
                    && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void LocalArrayOrFunctionCall_ResolvesToSmallestScopedDeclaration()
        {
            var code = @"
'note: Dim Foo() As Integer on this line would not compile in VBA
Public Sub DoSomething()
    Dim Foo() As Integer
    a = Foo(0) 'VBA raises index out of bounds error, i.e. VBA resolves to local Foo()
End Sub

Private Function Foo(ByVal bar As Integer)
    Foo = bar + 42
End Function";

            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable
                    && item.IsArray
                    && item.ParentScopeDeclaration.IdentifierName == "DoSomething");

                Assert.IsNotNull(declaration.References.SingleOrDefault(item => !item.IsAssignment));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void AnnotatedReference_LineAbove_HasAnnotations()
        {
            var code = @"
Public Sub DoSomething()
    Dim foo As Integer
    '@Ignore UnassignedVariableUsage
    a = foo
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && !item.IsUndeclared);

                var usage = declaration.References.Single();
                var annotation = (IgnoreAnnotation)usage.Annotations.First();
                Assert.IsTrue(
                    usage.Annotations.Count() == 1
                    && annotation.AnnotationType == AnnotationType.Ignore
                    && annotation.InspectionNames.Count() == 1
                    && annotation.InspectionNames.First() == "UnassignedVariableUsage");
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void AnnotatedReference_LinesAbove_HaveAnnotations()
        {
            var code = @"
Public Sub DoSomething()
    Dim foo As Integer
    '@Ignore UseMeaningfulName
    '@Ignore UnassignedVariableUsage
    a = foo
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && !item.IsUndeclared);

                var usage = declaration.References.Single();

                var annotation1 = (IgnoreAnnotation)usage.Annotations.ElementAt(0);
                var annotation2 = (IgnoreAnnotation)usage.Annotations.ElementAt(1);

                Assert.AreEqual(2, usage.Annotations.Count());
                Assert.AreEqual(AnnotationType.Ignore, annotation1.AnnotationType);
                Assert.AreEqual(AnnotationType.Ignore, annotation2.AnnotationType);

                Assert.IsTrue(usage.Annotations.Any(a => ((IgnoreAnnotation)a).InspectionNames.First() == "UseMeaningfulName"));
                Assert.IsTrue(usage.Annotations.Any(a => ((IgnoreAnnotation)a).InspectionNames.First() == "UnassignedVariableUsage"));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void AnnotatedDeclaration_LinesAbove_HaveAnnotations()
        {
            var code =
                @"'@TestMethod
'@IgnoreTest
Public Sub Foo()
End Sub";


            using (var state = Resolve(code))
            {
                var declaration = state.AllUserDeclarations.First(f => f.DeclarationType == DeclarationType.Procedure);

                Assert.IsTrue(declaration.Annotations.Count() == 2);
                Assert.IsTrue(declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.TestMethod));
                Assert.IsTrue(declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.IgnoreTest));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void AnnotatedReference_SameLine_HasNoAnnotations()
        {
            var code = @"
Public Sub DoSomething()
    Dim foo As Integer
    a = foo '@Ignore UnassignedVariableUsage 
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && !item.IsUndeclared);

                var usage = declaration.References.Single();

                Assert.IsTrue(!usage.Annotations.Any());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenUDT_NamedAfterProject_LocalResolvesToUDT()
        {
            var code = @"
Private Type TestProject1
    Foo As Integer
    Bar As String
End Type

Public Sub DoSomething()
    Dim Foo As TestProject1
    Foo.Bar = ""DoSomething""
    Foo.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedType);

                if (declaration.Project.Name != declaration.IdentifierName)
                {
                    Assert.Inconclusive("UDT should be named after project.");
                }

                var usage = declaration.References.SingleOrDefault();

                Assert.IsNotNull(usage);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenUDT_NamedAfterProject_FieldResolvesToUDT_EvenIfHiddenByLocal()
        {
            var code = @"
Private Type TestProject1
    Foo As Integer
    Bar As String
End Type

Private Foo As TestProject1

Public Sub DoSomething()
    Dim Foo As TestProject1
    Foo.Bar = ""DoSomething""
    Foo.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedType);

                if (declaration.Project.Name != declaration.IdentifierName)
                {
                    Assert.Inconclusive("UDT should be named after project.");
                }

                var usages = declaration.References;

                Assert.AreEqual(2, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenLocalVariable_NamedAfterUDTMember_ResolvesToLocalVariable()
        {
            var code = @"
Private Type TestProject1
    Foo As Integer
    Bar As String
End Type

Public Sub DoSomething()
    Dim Foo As TestProject1
    Foo.Bar = ""DoSomething""
    Foo.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable);

                if (declaration.Project.Name != declaration.AsTypeName)
                {
                    Assert.Inconclusive("variable should be named after project.");
                }
                var usages = declaration.References;

                Assert.AreEqual(2, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenLocalVariable_NamedAfterUDTMember_MemberCallResolvesToUDTMember()
        {
            var code = @"
Private Type TestProject1
    Foo As Integer
    Bar As String
End Type

Public Sub DoSomething()
    Dim Foo As TestProject1
    Foo.Bar = ""DoSomething""
    Foo.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedTypeMember
                    && item.IdentifierName == "Foo");

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenUDTMember_OfUDTType_ResolvesToDeclaredUDT()
        {
            var code = @"
Private Type TestProject1
    Foo As Integer
    Bar As String
End Type

Private Type Foo
    Foo As TestProject1
End Type

Public Sub DoSomething()
    Dim Foo As Foo
    Foo.Foo.Bar = ""DoSomething""
    Foo.Foo.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedTypeMember
                    && item.IdentifierName == "Foo"
                    && item.AsTypeName == item.Project.Name
                    && item.IdentifierName == item.ParentDeclaration.IdentifierName);

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(2, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenUDT_NamedAfterModule_LocalAsTypeResolvesToUDT()
        {
            var code = @"
Private Type TestProject1
    Foo As Integer
    Bar As String
End Type

Private Type TestModule1
    Foo As TestProject1
End Type

Public Sub DoSomething()
    Dim Foo As TestModule1
    Foo.Foo.Bar = ""DoSomething""
    Foo.Foo.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedType
                    && item.IdentifierName == item.ComponentName);

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenUDTMember_NamedAfterUDTType_NamedAfterModule_LocalAsTypeResolvesToUDT()
        {
            var code = @"
Private Type TestProject1
    Foo As Integer
    Bar As String
End Type

Private Type TestModule1
    TestModule1 As TestProject1
End Type

Public Sub DoSomething()
    Dim TestModule1 As TestModule1
    TestModule1.TestModule1.Bar = ""DoSomething""
    TestModule1.TestModule1.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedType
                    && item.IdentifierName == item.ComponentName);

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenField_NamedUnambiguously_FieldAssignmentCallResolvesToFieldDeclaration()
        {
            var code = @"
Private Type TestModule1
    Foo As Integer
End Type

Private Bar As TestModule1

Public Sub DoSomething()
    Bar.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable
                    && item.IdentifierName == "Bar");

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenField_NamedUnambiguously_InStatementFieldCallResolvesToFieldDeclaration()
        {
            var code = @"
Private Type TestModule1
    Foo As Integer
End Type

Private Bar As TestModule1

Public Sub DoSomething()
    a = Bar.Foo
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable
                    && item.IdentifierName == "Bar");

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenField_NamedAmbiguously_FieldAssignmentCallResolvesToFieldDeclaration()
        {
            var code = @"
Private Type TestModule1
    Foo As Integer
End Type

Private TestModule1 As TestModule1

Public Sub DoSomething()
    TestModule1.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable
                    && item.IdentifierName == item.ComponentName);

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenUDTField_NamedAmbiguously_MemberAssignmentCallResolvesToUDTMember()
        {
            var code = @"
Private Type TestModule1
    Foo As Integer
End Type

Private TestModule1 As TestModule1

Public Sub DoSomething()
    TestModule1.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedTypeMember
                    && item.IdentifierName == "Foo");

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenFullyReferencedUDTFieldMemberCall_ProjectParentMember_ResolvesToProject()
        {
            var code = @"
Private Type TestModule1
    Foo As Integer
End Type

Private TestModule1 As TestModule1

Public Sub DoSomething()
    TestProject1.TestModule1.TestModule1.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Project
                    && item.IdentifierName == item.Project.Name);

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenFullyQualifiedUDTFieldMemberCall_ModuleParentMember_ResolvesToModule()
        {
            var code = @"
Private Type TestModule1
    Foo As Integer
End Type

Private TestModule1 As TestModule1

Public Sub DoSomething()
    TestProject1.TestModule1.TestModule1.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.ProceduralModule
                    && item.IdentifierName == item.ComponentName);

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenFullyQualifiedUDTFieldMemberCall_FieldParentMember_ResolvesToVariable()
        {
            var code = @"
Private Type TestModule1
    Foo As Integer
End Type

Private TestModule1 As TestModule1

Public Sub DoSomething()
    TestProject1.TestModule1.TestModule1.Foo = 42
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable
                    && item.IdentifierName == item.ComponentName);

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenGlobalVariable_QualifiedUsageInOtherModule_AssignmentCallResolvesToVariable()
        {
            var code_module1 = @"
Private Type TSomething
    Foo As Integer
    Bar As Integer
End Type

Public Something As TSomething
";

            var code_module2 = @"
Sub DoSomething()
    Component1.Something.Bar = 42
End Sub";

            var module1 = Tuple.Create(code_module1, ComponentType.StandardModule);
            var module2 = Tuple.Create(code_module2, ComponentType.StandardModule);
            using (var state = Resolve(module1, module2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable
                    && item.Accessibility == Accessibility.Public
                    && item.IdentifierName == "Something");

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenGlobalVariable_QualifiedUsageInOtherModule_CallResolvesToVariable()
        {
            var code_module1 = @"
Private Type TSomething
    Foo As Integer
    Bar As Integer
End Type

Public Something As TSomething
";

            var code_module2 = @"
Sub DoSomething()
    a = Component1.Something.Bar
End Sub
";

            var module1 = Tuple.Create(code_module1, ComponentType.StandardModule);
            var module2 = Tuple.Create(code_module2, ComponentType.StandardModule);
            using (var state = Resolve(module1, module2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable
                    && item.Accessibility == Accessibility.Public
                    && item.IdentifierName == "Something");

                var usages = declaration.References.Where(item =>
                    item.ParentScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void RedimStmt_RedimVariableDeclarationIsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced() As Variant
    ReDim referenced(referenced TO referenced, referenced), referenced(referenced)
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(6, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void OpenStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Open referenced For Binary Access Read Lock Read As #referenced Len = referenced
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(3, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void CloseStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Close referenced, referenced
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void SeekStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Seek #referenced, referenced
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void LockStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Lock referenced, referenced To referenced
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(3, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void UnlockStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Unlock referenced, referenced To referenced
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(3, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void LineInputStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Line Input #referenced, referenced
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void LineInputStmt_ReferenceIsAssignment()
        {
            var code = @"
Public Sub Test()
    Dim file As Integer, content
    Line Input #file, content
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "content");

                Assert.IsTrue(declaration.References.Single().IsAssignment);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void WidthStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Width #referenced, referenced
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void PrintStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Print #referenced,,referenced; SPC(referenced), TAB(referenced)
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(4, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void WriteStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Write #referenced,,referenced; SPC(referenced), TAB(referenced)
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(4, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void InputStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Input #referenced,referenced
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(2, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void InputStmt_ReferenceIsAssignment()
        {
            var code = @"
Public Sub Test()
    Dim str As String
    Dim xCoord, yCoord, zCoord As Double
    Input #1, str, xCoord, yCoord, zCoord
End Sub
";
            using (var state = Resolve(code))
            {

                var strDeclaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "str");

                var xCoordDeclaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "xCoord");

                var yCoordDeclaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "yCoord");

                var zCoordDeclaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "zCoord");

                Assert.IsTrue(strDeclaration.References.Single().IsAssignment);
                Assert.IsTrue(xCoordDeclaration.References.Single().IsAssignment);
                Assert.IsTrue(yCoordDeclaration.References.Single().IsAssignment);
                Assert.IsTrue(zCoordDeclaration.References.Single().IsAssignment);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void PutStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Put referenced,referenced,referenced
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(3, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GetStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Get #referenced,referenced,referenced
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(3, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GetStmt_ReferenceIsAssignment()
        {
            var code = @"
Public Sub Test()
    Dim fileNumber As Integer, recordNumber, variable
    Get #fileNumber, recordNumber, variable
End Sub
";
            using (var state = Resolve(code))
            {

                var variableDeclaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "variable");

                Assert.IsTrue(variableDeclaration.References.Single().IsAssignment);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void LineSpecialForm_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Me.Line (referenced, referenced)-(referenced, referenced)
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(4, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void CircleSpecialForm_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Me.Circle Step(referenced, referenced), referenced, referenced, referenced, referenced, referenced
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(7, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ScaleSpecialForm_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Scale (referenced, referenced)-(referenced, referenced)
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

                Assert.AreEqual(4, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void FieldLengthStmt_IsReferenceToLocalVariable()
        {
            var code = @"
Public Sub Test()
    Const Len As Integer = 4
    Dim a As String * Len
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Constant && item.IdentifierName == "Len");

                Assert.AreEqual(1, declaration.References.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenControlDeclaration_ResolvesUsageInCodeBehind()
        {
            var code = @"
Public Sub DoSomething()
    TextBox1.Height = 20
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected);
            var form = project.MockUserFormBuilder("Form1", code).AddControl("TextBox1").Build();
            project.AddComponent(form);
            builder.AddProject(project.Build());
            var vbe = builder.Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                if (state.Status == ParserState.ResolverError)
                {
                    Assert.Fail("Parser state should be 'Ready', but returns '{0}'.", state.Status);
                }
                if (state.Status != ParserState.Ready)
                {
                    Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", state.Status);
                }

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Control
                    && item.IdentifierName == "TextBox1");

                var usages = declaration.References.Where(item =>
                    item.ParentNonScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenLocalDeclarationAsQualifiedClassName_ResolvesFirstPartToProject()
        {
            var code_class1 = @"
Public Sub DoSomething
    Dim foo As TestProject1.Class2
End Sub
";
            var code_class2 = @"
Public Type TFoo
    Bar As Integer
End Type

Private this As TFoo

Public Property Get Bar() As Integer
    Bar = this.Bar
End Property
";
            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Project
                    && item.IdentifierName == "TestProject1");

                var usages = declaration.References.Where(item =>
                    item.ParentNonScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(state.Status, ParserState.Ready);
                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenLocalDeclarationAsQualifiedClassName_ResolvesSecondPartToClassModule()
        {
            var code_class1 = @"
Public Sub DoSomething
    Dim foo As TestProject1.Class2
End Sub
";
            var code_class2 = @"
Public Type TFoo
    Bar As Integer
End Type

Private this As TFoo

Public Property Get Bar() As Integer
    Bar = this.Bar
End Property
";
            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.ClassModule
                    && item.IdentifierName == "Class2");

                var usages = declaration.References.Where(item =>
                    item.ParentNonScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void GivenLocalDeclarationAsQualifiedClassName_ResolvesThirdPartToUDT()
        {
            var code_class1 = @"
Public Sub DoSomething
    Dim foo As TestProject1.Class2.TFoo
End Sub
";
            var code_class2 = @"
Public Type TFoo
    Bar As Integer
End Type

Private this As TFoo

Public Property Get Bar() As Integer
    Bar = this.Bar
End Property
";
            using (var state = Resolve(code_class1, code_class2))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.UserDefinedType
                    && item.IdentifierName == "TFoo");

                var usages = declaration.References.Where(item =>
                    item.ParentNonScoping.IdentifierName == "DoSomething");

                Assert.AreEqual(1, usages.Count());
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void QualifiedSetStatement_FirstSectionDoesNotHaveAssignmentFlag()
        {
            var variableDeclarationClass = @"
Public foo As Boolean
";

            var classVariableDeclarationClass = @"
Public myClass As Class1
";

            var variableCallClass = @"
Public Sub bar()
    Dim myClassN As Class2
    Set myClassN.myClass.foo = True
End Sub
";
            using (var state = Resolve(variableDeclarationClass, classVariableDeclarationClass, variableCallClass))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "myClassN");

                Assert.IsFalse(declaration.References.ElementAt(0).IsAssignment);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void QualifiedSetStatement_MiddleSectionDoesNotHaveAssignmentFlag()
        {
            var variableDeclarationClass = @"
Public foo As Boolean
";

            var classVariableDeclarationClass = @"
Public myClass As Class1
";

            var variableCallClass = @"
Public Sub bar()
    Dim myClassN As Class2
    Set myClassN.myClass.foo = True
End Sub
";
            using (var state = Resolve(variableDeclarationClass, classVariableDeclarationClass, variableCallClass))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "myClass");

                Assert.IsFalse(declaration.References.ElementAt(0).IsAssignment);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void QualifiedSetStatement_LastSectionHasAssignmentFlag()
        {
            var variableDeclarationClass = @"
Public foo As Boolean
";

            var classVariableDeclarationClass = @"
Public myClass As Class1
";

            var variableCallClass = @"
Public Sub bar()
    Dim myClassN As Class2
    Set myClassN.myClass.foo = True
End Sub
";
            using (var state = Resolve(variableDeclarationClass, classVariableDeclarationClass, variableCallClass))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

                Assert.IsTrue(declaration.References.ElementAt(0).IsAssignment);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void SetStatement_HasAssignmentFlag()
        {
            var variableDeclarationClass = @"
Public foo As Variant

Public Sub bar()
    Set foo = New Class2
End Sub
";
            using (var state = Resolve(variableDeclarationClass, string.Empty))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

                Assert.IsTrue(declaration.References.ElementAt(0).IsAssignment);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ImplicitLetStatement_HasAssignmentFlag()
        {
            var variableDeclarationClass = @"
Public foo As Boolean

Public Sub bar()
    foo = True
End Sub
";
            using (var state = Resolve(variableDeclarationClass))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

                Assert.IsTrue(declaration.References.ElementAt(0).IsAssignment);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void ExplicitLetStatement_HasAssignmentFlag()
        {
            var variableDeclarationClass = @"
Public foo As Boolean

Public Sub bar()
    Let foo = True
End Sub
";
            using (var state = Resolve(variableDeclarationClass))
            {

                var declaration = state.AllUserDeclarations.Single(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

                Assert.IsTrue(declaration.References.ElementAt(0).IsAssignment);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        //https://github.com/rubberduck-vba/Rubberduck/issues/2478
        [Test]
        public void VariableNamedBfResolvesAsAVariable()
        {
            var code = @"
Public Sub Test()
    Dim bf As Integer
    Debug.Print bf
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.SingleOrDefault(item =>
                    item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "bf");

                Assert.IsNotNull(declaration);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        //https://github.com/rubberduck-vba/Rubberduck/issues/2478
        [Test]
        public void ProcedureNamedBfResolvesAsAProcedure()
        {
            var code = @"
Public Sub Bf()
    Debug.Print ""I'm cool with that""
End Sub
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.SingleOrDefault(item =>
                    item.DeclarationType == DeclarationType.Procedure && item.IdentifierName == "Bf");

                Assert.IsNotNull(declaration);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        //https://github.com/rubberduck-vba/Rubberduck/issues/2478
        [Test]
        public void TypeNamedBfResolvesAsAType()
        {
            var code = @"
Private Type Bf
    b As Long
    f As Long
End Type
";
            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.SingleOrDefault(item =>
                    item.DeclarationType == DeclarationType.UserDefinedType && item.IdentifierName == "Bf");

                Assert.IsNotNull(declaration);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        //https://github.com/rubberduck-vba/Rubberduck/issues/2523
        [Test]
        public void AnnotationFollowedByCommentAnnotatesDeclaration()
        {
            var code = @"
Public Sub DoSomething()
    '@Ignore VariableNotAssigned: that's actually a Rubberduck bug, see #2522
    ReDim orgs(0 To items.Count - 1, 0 To 1)
End Sub
";
            var results = new[] { "VariableNotAssigned" };

            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item => item.IdentifierName == "orgs");

                var annotation = declaration.Annotations.SingleOrDefault(item => item.AnnotationType == AnnotationType.Ignore);
                Assert.IsNotNull(annotation);
                Assert.IsTrue(results.SequenceEqual(((IgnoreAnnotation)annotation).InspectionNames));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        //https://github.com/rubberduck-vba/Rubberduck/issues/2523
        [Test]
        public void AnnotationListFollowedByCommentAnnotatesDeclaration()
        {
            var code = @"
Public Sub DoSomething()
    '@Ignore VariableNotAssigned, UndeclaredVariable, VariableNotUsed: Ignore ALL the inspections!
    ReDim orgs(0 To items.Count - 1, 0 To 1)
End Sub
";

            var results = new[] { "VariableNotAssigned", "UndeclaredVariable", "VariableNotUsed" };

            using (var state = Resolve(code))
            {

                var declaration = state.AllUserDeclarations.Single(item => item.IdentifierName == "orgs");

                var annotation = declaration.Annotations.SingleOrDefault(item => item.AnnotationType == AnnotationType.Ignore);
                Assert.IsNotNull(annotation);
                Assert.IsTrue(results.SequenceEqual(((IgnoreAnnotation)annotation).InspectionNames));
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void MemberReferenceIsAssignmentTarget_NotTheParentObject()
        {
            var class1 = @"
Option Explicit
Private mSomething As Long
Public Property Get Something() As Long
    Something = mSomething
End Property
Public Property Let Something(ByVal value As Long)
    mSomething = value
End Property
";
            var caller = @"
Option Explicit
Private Sub DoSomething(ByVal foo As Class1)
    foo.Something = 42
End Sub
";
            using (var state = Resolve(class1, caller))
            {

                var declaration = state.AllUserDeclarations.Single(item => item.IdentifierName == "foo" && item.DeclarationType == DeclarationType.Parameter);
                var reference = declaration.References.Single();

                Assert.IsFalse(reference.IsAssignment, "LHS member call on object is treating the object itself as an assignment target.");

                var member = state.AllUserDeclarations.Single(item => item.IdentifierName == "Something" && item.DeclarationType == DeclarationType.PropertyLet);
                var call = member.References.Single();

                Assert.IsTrue(call.IsAssignment, "LHS member call on object is not flagging member reference as assignment target.");
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void AttributeFollowingSubGetsAssignedToSub()
        {
            var code = @"
Public Sub Foo(): End Sub
Attribute Foo.VB_Description = ""Foo description""

Public Sub Bar()
End Sub
";
            using (var state = Resolve(code))
            {
                var declaration = state.DeclarationFinder.MatchName("Foo").Single(item => item.DeclarationType == DeclarationType.Procedure);

                var expectedDescription = "Foo description";
                var actualDescription = declaration.DescriptionString;
                Assert.AreEqual(expectedDescription, actualDescription);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void AttributeFollowingFunctionGetsAssignedToFunction()
        {
            var code = @"
Public Function Foo(): End Function
Attribute Foo.VB_Description = ""Foo description""

Public Sub Bar()
End Sub
";
            using (var state = Resolve(code))
            {
                var declaration = state.DeclarationFinder.MatchName("Foo").Single(item => item.DeclarationType == DeclarationType.Function);

                var expectedDescription = "Foo description";
                var actualDescription = declaration.DescriptionString;
                Assert.AreEqual(expectedDescription, actualDescription);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void AttributeFollowingPropertyGetGetsAssignedToPropertyGet()
        {
            var code = @"
Public Property Get Foo(): End Property
Attribute Foo.VB_Description = ""Foo description""

Public Sub Bar()
End Sub
";
            using (var state = Resolve(code))
            {
                var declaration = state.DeclarationFinder.MatchName("Foo").Single(item => item.DeclarationType == DeclarationType.PropertyGet);

                var expectedDescription = "Foo description";
                var actualDescription = declaration.DescriptionString;
                Assert.AreEqual(expectedDescription, actualDescription);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void AttributeFollowingPropertyLetGetsAssignedToPropertyLet()
        {
            var code = @"
Public Property Let Foo(): End Property
Attribute Foo.VB_Description = ""Foo description""

Public Sub Bar()
End Sub
";
            using (var state = Resolve(code))
            {
                var declaration = state.DeclarationFinder.MatchName("Foo").Single(item => item.DeclarationType == DeclarationType.PropertyLet);

                var expectedDescription = "Foo description";
                var actualDescription = declaration.DescriptionString;
                Assert.AreEqual(expectedDescription, actualDescription);
            }
        }

        [Category("Grammar")]
        [Category("Resolver")]
        [Test]
        public void AttributeFollowingPropertySetGetsAssignedToPropertySet()
        {
            var code = @"
Public Property Set Foo(): End Property
Attribute Foo.VB_Description = ""Foo description""

Public Sub Bar()
End Sub
";
            using (var state = Resolve(code))
            {
                var declaration = state.DeclarationFinder.MatchName("Foo").Single(item => item.DeclarationType == DeclarationType.PropertySet);

                var expectedDescription = "Foo description";
                var actualDescription = declaration.DescriptionString;
                Assert.AreEqual(expectedDescription, actualDescription);
            }
        }
    }
}

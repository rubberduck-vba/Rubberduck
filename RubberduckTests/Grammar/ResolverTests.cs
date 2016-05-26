using System;
using System.Linq;
using Microsoft.Vbe.Interop;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using Rubberduck.Parsing.Annotations;

namespace RubberduckTests.Grammar
{
    [TestClass]
    public class ResolverTests
    {
        private RubberduckParserState Resolve(string code, vbext_ComponentType moduleType = vbext_ComponentType.vbext_ct_StdModule)
        {
            var builder = new MockVbeBuilder();
            VBComponent component;
            var vbe = builder.BuildFromSingleModule(code, moduleType, out component, new Rubberduck.VBEditor.Selection());
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.ResolverError)
            {
                Assert.Fail("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }

            return parser.State;
        }

        private RubberduckParserState Resolve(params string[] classes)
        {
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none);
            for (var i = 0; i < classes.Length; i++)
            {
                projectBuilder.AddComponent("Class" + (i + 1), vbext_ComponentType.vbext_ct_ClassModule, classes[i]);
            }

            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.ResolverError)
            {
                Assert.Fail("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }

            return parser.State;
        }

        private RubberduckParserState Resolve(params Tuple<string, vbext_ComponentType>[] components)
        {
            var builder = new MockVbeBuilder();
            var projectBuilder = builder.ProjectBuilder("TestProject", vbext_ProjectProtection.vbext_pp_none);
            for (var i = 0; i < components.Length; i++)
            {
                projectBuilder.AddComponent("Component" + (i + 1), components[i].Item2, components[i].Item1);
            }

            var project = projectBuilder.Build();
            builder.AddProject(project);
            var vbe = builder.Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.ResolverError)
            {
                Assert.Fail("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }

            return parser.State;
        }

        [TestMethod]
        public void FunctionReturnValueAssignment_IsReferenceToFunctionDeclaration()
        {
            // arrange
            var code = @"
Public Function Foo() As String
    Foo = 42
End Function
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Function && item.IdentifierName == "Foo");

            Assert.AreEqual(1, declaration.References.Count(item => item.IsAssignment));
        }

        [TestMethod]
        public void TypeOfIsExpression_BooleanExpressionIsReferenceToLocalVariable()
        {
            // arrange
            var code_class1 = @"
Public Function Foo() As String
    Dim a As Object
    anything = TypeOf a Is Class2
End Function
";
            // We only use the second class as as target of the type expression, its contents don't matter.
            var code_class2 = string.Empty;

            // act
            var state = Resolve(code_class1, code_class2);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "a");

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void TypeOfIsExpression_TypeExpressionIsReferenceToClass()
        {
            // arrange
            var code_class1 = @"
Public Function Foo() As String
    Dim a As Object
    anything = TypeOf a Is Class2
End Function
";
            // We only use the second class as as target of the type expression, its contents don't matter.
            var code_class2 = string.Empty;

            // act
            var state = Resolve(code_class1, code_class2);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.ClassModule && item.IdentifierName == "Class2");

            Assert.AreEqual(1, declaration.References.Count());
        }

        [TestMethod]
        public void FunctionCall_IsReferenceToFunctionDeclaration()
        {
            // arrange
            var code = @"
Public Sub DoSomething()
    Foo
End Sub

Private Function Foo() As String
    Foo = 42
End Function
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Function && item.IdentifierName == "Foo");

            var reference = declaration.References.SingleOrDefault(item => !item.IsAssignment);
            Assert.IsNotNull(reference);
            Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
        }

        [TestMethod]
        public void LocalVariableCall_IsReferenceToVariableDeclaration()
        {
            // arrange
            var code = @"
Public Sub DoSomething()
    Dim foo As Integer
    a = foo
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

            var reference = declaration.References.SingleOrDefault(item => !item.IsAssignment);
            Assert.IsNotNull(reference);
            Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
        }

        [TestMethod]
        public void LocalVariableForeignNameCall_IsReferenceToVariableDeclaration()
        {
            // arrange
            var code = @"
Public Sub DoSomething()
    Dim foo As Integer
    a = [foo]
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

            var reference = declaration.References.SingleOrDefault(item => !item.IsAssignment);
            Assert.IsNotNull(reference);
            Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
        }

        [TestMethod]
        public void LocalVariableAssignment_IsReferenceToVariableDeclaration()
        {
            // arrange
            var code = @"
Public Sub DoSomething()
    Dim foo As Integer
    foo = 42
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

            var reference = declaration.References.SingleOrDefault(item => item.IsAssignment);
            Assert.IsNotNull(reference);
            Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
        }

        [TestMethod]
        public void SingleLineIfStatementLabel_IsReferenceToLabel()
        {
            // arrange
            var code = @"
Public Sub DoSomething()
    If True Then 5
5:
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.LineLabel && item.IdentifierName == "5");

            var reference = declaration.References.SingleOrDefault();
            Assert.IsNotNull(reference);
            Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
        }

        [TestMethod]
        public void ProjectUdtSameNameFirstProjectThenUdt_FirstReferenceIsToProject()
        {
            // arrange
            var code = string.Format(@"
Private Type {0}
    anything As String
End Type

Public Sub DoSomething()
    Dim a As {0}.{1}.{0}
End Sub
", MockVbeBuilder.TEST_PROJECT_NAME, MockVbeBuilder.TEST_MODULE_NAME);
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Project && item.IdentifierName == MockVbeBuilder.TEST_PROJECT_NAME);

            var reference = declaration.References.SingleOrDefault();
            Assert.IsNotNull(reference);
            Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
        }

        [TestMethod]
        public void ProjectUdtSameNameUdtOnly_IsReferenceToUdt()
        {
            // arrange
            var code = string.Format(@"
Private Type {0}
    anything As String
End Type

Public Sub DoSomething()
    Dim a As {1}.{0}
End Sub
", MockVbeBuilder.TEST_PROJECT_NAME, MockVbeBuilder.TEST_MODULE_NAME);
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedType && item.IdentifierName == MockVbeBuilder.TEST_PROJECT_NAME);

            var reference = declaration.References.SingleOrDefault();
            Assert.IsNotNull(reference);
            Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
        }

        [TestMethod]
        public void EncapsulatedVariableAssignment_DoesNotResolve()
        {
            // arrange
            var code_class1 = @"
Public Sub DoSomething()
    foo = 42
End Sub
";
            var code_class2 = @"
Option Explicit
Public foo As Integer
";
            var class1 = Tuple.Create(code_class1, vbext_ComponentType.vbext_ct_ClassModule);
            var class2 = Tuple.Create(code_class2, vbext_ComponentType.vbext_ct_ClassModule);

            // act
            var state = Resolve(class1, class2);

            // assert
            var declaration = state.AllUserDeclarations.Single(item => item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

            var reference = declaration.References.SingleOrDefault(item => item.IsAssignment);
            Assert.IsNull(reference);
        }

        [TestMethod]
        public void PublicVariableCall_IsReferenceToVariableDeclaration()
        {
            // arrange
            var code_class1 = @"
Public Sub DoSomething()
    a = foo
End Sub
";
            var code_class2 = @"
Option Explicit
Public foo As Integer
";
            // act
            var state = Resolve(
                Tuple.Create(code_class1, vbext_ComponentType.vbext_ct_ClassModule),
                Tuple.Create(code_class2, vbext_ComponentType.vbext_ct_StdModule));

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

            var reference = declaration.References.SingleOrDefault(item => !item.IsAssignment);
            Assert.IsNotNull(reference);
            Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
        }

        [TestMethod]
        public void PublicVariableAssignment_IsReferenceToVariableDeclaration()
        {
            // arrange
            var code_class1 = @"
Public Sub DoSomething()
    foo = 42
End Sub
";
            var code_module1 = @"
Option Explicit
Public foo As Integer
";
            var class1 = Tuple.Create(code_class1, vbext_ComponentType.vbext_ct_ClassModule);
            var module1 = Tuple.Create(code_module1, vbext_ComponentType.vbext_ct_StdModule);

            // act
            var state = Resolve(class1, module1);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

            var reference = declaration.References.SingleOrDefault(item => item.IsAssignment);
            Assert.IsNotNull(reference);
            Assert.AreEqual("DoSomething", reference.ParentScoping.IdentifierName);
        }

        [TestMethod]
        public void UserDefinedTypeVariableAsTypeClause_IsReferenceToUserDefinedTypeDeclaration()
        {
            // arrange
            var code = @"
Private Type TFoo
    Bar As Integer
End Type
Private this As TFoo
";
            // act
            var state = Resolve(code, vbext_ComponentType.vbext_ct_ClassModule);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedType && item.IdentifierName == "TFoo");

            Assert.IsNotNull(declaration.References.SingleOrDefault());
        }

        [TestMethod]
        public void ObjectVariableAsTypeClause_IsReferenceToClassModuleDeclaration()
        {
            // arrange
            var code_class1 = @"
Public Sub DoSomething()
    Dim foo As Class2
End Sub
";
            var code_class2 = @"
Option Explicit
";

            // act
            var state = Resolve(code_class1, code_class2);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.ClassModule && item.IdentifierName == "Class2");

            Assert.IsNotNull(declaration.References.SingleOrDefault());
        }

        [TestMethod]
        public void ParameterCall_IsReferenceToParameterDeclaration()
        {
            // arrange
            var code = @"
Public Sub DoSomething(ByVal foo As Integer)
    a = foo
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Parameter && item.IdentifierName == "foo");

            Assert.IsNotNull(declaration.References.SingleOrDefault());
        }

        [TestMethod]
        public void ParameterAssignment_IsAssignmentReferenceToParameterDeclaration()
        {
            // arrange
            var code = @"
Public Sub DoSomething(ByRef foo As Integer)
    foo = 42
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Parameter && item.IdentifierName == "foo");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item => item.IsAssignment));
        }

        [TestMethod]
        public void NamedParameterCall_IsReferenceToParameterDeclaration()
        {
            // arrange
            var code = @"
Public Sub DoSomething()
    DoSomethingElse foo:=42
End Sub

Private Sub DoSomethingElse(ByVal foo As Integer)
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Parameter && item.IdentifierName == "foo");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.IdentifierName == "DoSomething"));
        }

        [TestMethod]
        public void UserDefinedTypeMemberCall_IsReferenceToUserDefinedTypeMemberDeclaration()
        {
            // arrange
            var code = @"
Private Type TFoo
    Bar As Integer
End Type
Private this As TFoo

Public Property Get Bar() As Integer
    Bar = this.Bar
End Property
";
            // act
            var state = Resolve(code, vbext_ComponentType.vbext_ct_ClassModule);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedTypeMember && item.IdentifierName == "Bar");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.DeclarationType == DeclarationType.PropertyGet
                && item.ParentScoping.IdentifierName == "Bar"));
        }

        [TestMethod]
        public void UserDefinedTypeVariableCall_IsReferenceToVariableDeclaration()
        {
            // arrange
            var code = @"
Private Type TFoo
    Bar As Integer
End Type
Private this As TFoo

Public Property Get Bar() As Integer
    Bar = this.Bar
End Property
";
            // act
            var state = Resolve(code, vbext_ComponentType.vbext_ct_ClassModule);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "this");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.DeclarationType == DeclarationType.PropertyGet
                && item.ParentScoping.IdentifierName == "Bar"));
        }

        [TestMethod]
        public void WithVariableMemberCall_IsReferenceToMemberDeclaration()
        {
            // arrange
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
            // act
            var state = Resolve(code_class1, code_class2);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.PropertyGet && item.IdentifierName == "Foo");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.DeclarationType == DeclarationType.Procedure
                && item.ParentScoping.IdentifierName == "DoSomething"));
        }

        [TestMethod]
        public void NestedWithVariableMemberCall_IsReferenceToMemberDeclaration()
        {
            // arrange
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

            // act
            var state = Resolve(code_class1, code_class2, code_class3);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.PropertyGet && item.IdentifierName == "Bar");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.DeclarationType == DeclarationType.Procedure
                && item.ParentScoping.IdentifierName == "DoSomething"));
        }

        [TestMethod]
        public void ResolvesLocalVariableToSmallestScopeIdentifier()
        {
            var code = @"
Private foo As Integer

Private Sub DoSomething()
    Dim foo As Integer
    foo = 42
End Sub
";
            // act
            var state = Resolve(code);

            // assert
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

        [TestMethod]
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
            // act
            var state = Resolve(code_class1, code_class2);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.ClassModule && item.IdentifierName == "Class1");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.IdentifierName == "Class2"));
        }

        [TestMethod]
        public void NestedMemberCall_IsReferenceToMember()
        {
            // arrange
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
            // act
            var state = Resolve(code_class1, code_class2, code_class3);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.PropertyGet && item.IdentifierName == "Bar");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.DeclarationType == DeclarationType.Procedure
                && item.ParentScoping.IdentifierName == "DoSomething"));
        }

        [TestMethod]
        public void MemberCallParent_IsReferenceToParent()
        {
            // arrange
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
            // act
            var state = Resolve(code_class1, code_class2);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Parameter && item.IdentifierName == "a");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.DeclarationType == DeclarationType.Procedure
                && item.ParentScoping.IdentifierName == "DoSomething"));
        }

        [TestMethod]
        public void ForLoop_IsAssignmentReferenceToIteratorDeclaration()
        {
            // arrange
            var code = @"
Public Sub DoSomething()
    Dim i As Integer
    For i = 0 To 9
    Next
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "i");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.DeclarationType == DeclarationType.Procedure
                && item.ParentScoping.IdentifierName == "DoSomething"
                && item.IsAssignment));
        }

        [TestMethod]
        public void ForEachLoop_IsReferenceToIteratorDeclaration()
        {
            var code = @"
Public Sub DoSomething(ByVal c As Collection)
    Dim i As Variant
    For Each i In c
    Next
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "i");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.DeclarationType == DeclarationType.Procedure
                && item.ParentScoping.IdentifierName == "DoSomething"
                && item.IsAssignment));
        }

        [TestMethod]
        public void ForEachLoop_InClauseIsReferenceToIteratedDeclaration()
        {
            var code = @"
Public Sub DoSomething(ByVal c As Collection)
    Dim i As Variant
    For Each i In c
    Next
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Parameter && item.IdentifierName == "c");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.DeclarationType == DeclarationType.Procedure
                && item.ParentScoping.IdentifierName == "DoSomething"
                && !item.IsAssignment));
        }

        [TestMethod]
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
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Parameter
                && item.IdentifierName == "values"
                && item.IsArray);

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.DeclarationType == DeclarationType.Procedure
                && item.ParentScoping.IdentifierName == "DoSomething"
                && !item.IsAssignment));
        }

        [TestMethod]
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
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Parameter
                && item.IdentifierName == "values"
                && item.IsArray);

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.DeclarationType == DeclarationType.Procedure
                && item.ParentScoping.IdentifierName == "DoSomething"
                && item.IsAssignment));
        }

        [TestMethod]
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

            var state = Resolve(code_class1, code_class2);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.PropertyGet
                && item.IdentifierName == "Foo");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                !item.IsAssignment
                && item.ParentScoping.IdentifierName == "DoSomething"
                && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
        }

        [TestMethod]
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

            var state = Resolve(code_class1, code_class2);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.PropertySet
                && item.IdentifierName == "Foo");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.IsAssignment
                && item.ParentScoping.IdentifierName == "DoSomething"
                && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
        }

        [TestMethod]
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

            var state = Resolve(code_class1, code_class2);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.PropertyLet
                && item.IdentifierName == "Foo");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.IsAssignment
                && item.ParentScoping.IdentifierName == "DoSomething"
                && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.EnumerationMember
                && item.IdentifierName == "Foo"
                && item.ParentDeclaration.IdentifierName == "FooBarBaz");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                !item.IsAssignment
                && item.ParentScoping.IdentifierName == "DoSomething"
                && item.ParentScoping.DeclarationType == DeclarationType.Procedure));

        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.EnumerationMember
                && item.IdentifierName == "Foo"
                && item.ParentDeclaration.IdentifierName == "FooBarBaz");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                !item.IsAssignment
                && item.ParentScoping.IdentifierName == "DoSomething"
                && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
        }

        [TestMethod]
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

            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Enumeration
                && item.IdentifierName == "FooBarBaz");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.IdentifierName == "DoSomething"
                && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
        }

        [TestMethod]
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

            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedType
                && item.IdentifierName == "TFoo");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item =>
                item.ParentScoping.IdentifierName == "DoSomething"
                && item.ParentScoping.DeclarationType == DeclarationType.Procedure));
        }

        [TestMethod]
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

            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable
                && item.IsArray
                && item.ParentScopeDeclaration.IdentifierName == "DoSomething");

            Assert.IsNotNull(declaration.References.SingleOrDefault(item => !item.IsAssignment));
        }

        [TestMethod]
        public void AnnotatedReference_LineAbove_HasAnnotations()
        {
            var code = @"
Public Sub DoSomething()
    Dim foo As Integer
    '@Ignore UnassignedVariableUsage
    a = foo
End Sub
";
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable);

            var usage = declaration.References.Single();
            var annotation = (IgnoreAnnotation)usage.Annotations.First();
            Assert.IsTrue(
                usage.Annotations.Count() == 1
                && annotation.AnnotationType == AnnotationType.Ignore
                && annotation.InspectionNames.Count() == 1
                && annotation.InspectionNames.First() == "UnassignedVariableUsage");
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable);

            var usage = declaration.References.Single();

            var annotation1 = (IgnoreAnnotation)usage.Annotations.ElementAt(0);
            var annotation2 = (IgnoreAnnotation)usage.Annotations.ElementAt(1);

            Assert.AreEqual(2, usage.Annotations.Count());
            Assert.AreEqual(AnnotationType.Ignore, annotation1.AnnotationType);
            Assert.AreEqual(AnnotationType.Ignore, annotation2.AnnotationType);

            Assert.IsTrue(usage.Annotations.Any(a => ((IgnoreAnnotation)a).InspectionNames.First() == "UseMeaningfulName"));
            Assert.IsTrue(usage.Annotations.Any(a => ((IgnoreAnnotation)a).InspectionNames.First() == "UnassignedVariableUsage"));
        }

        [TestMethod]
        public void AnnotatedDeclaration_LinesAbove_HaveAnnotations()
        {
            var code =
@"'@TestMethod
'@IgnoreTest
Public Sub Foo()
End Sub";


            var state = Resolve(code);
            var declaration = state.AllUserDeclarations.First(f => f.DeclarationType == DeclarationType.Procedure);

            Assert.IsTrue(declaration.Annotations.Count() == 2);
            Assert.IsTrue(declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.TestMethod));
            Assert.IsTrue(declaration.Annotations.Any(a => a.AnnotationType == AnnotationType.IgnoreTest));
        }

        [TestMethod]
        public void AnnotatedReference_SameLine_HasNoAnnotations()
        {
            var code = @"
Public Sub DoSomething()
    Dim foo As Integer
    a = foo '@Ignore UnassignedVariableUsage 
End Sub
";
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable);

            var usage = declaration.References.Single();

            Assert.IsTrue(!usage.Annotations.Any());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedType);

            if (declaration.Project.Name != declaration.IdentifierName)
            {
                Assert.Inconclusive("UDT should be named after project.");
            }

            var usage = declaration.References.SingleOrDefault();

            Assert.IsNotNull(usage);
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedType);

            if (declaration.Project.Name != declaration.IdentifierName)
            {
                Assert.Inconclusive("UDT should be named after project.");
            }

            var usages = declaration.References;

            Assert.AreEqual(2, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable);

            if (declaration.Project.Name != declaration.AsTypeName)
            {
                Assert.Inconclusive("variable should be named after project.");
            }
            var usages = declaration.References;

            Assert.AreEqual(2, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedTypeMember
                && item.IdentifierName == "Foo");

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedTypeMember
                && item.IdentifierName == "Foo"
                && item.AsTypeName == item.Project.Name
                && item.IdentifierName == item.ParentDeclaration.IdentifierName);

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(2, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedType
                && item.IdentifierName == item.ComponentName);

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedType
                && item.IdentifierName == item.ComponentName);

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable
                && item.IdentifierName == "Bar");

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable
                && item.IdentifierName == "Bar");

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable
                && item.IdentifierName == item.ComponentName);

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedTypeMember
                && item.IdentifierName == "Foo");

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Project
                && item.IdentifierName == item.Project.Name);

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.ProceduralModule
                && item.IdentifierName == item.ComponentName);

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable
                && item.IdentifierName == item.ComponentName);

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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

            var module1 = Tuple.Create(code_module1, vbext_ComponentType.vbext_ct_StdModule);
            var module2 = Tuple.Create(code_module2, vbext_ComponentType.vbext_ct_StdModule);
            var state = Resolve(module1, module2);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable
                && item.Accessibility == Accessibility.Public
                && item.IdentifierName == "Something");

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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

            var module1 = Tuple.Create(code_module1, vbext_ComponentType.vbext_ct_StdModule);
            var module2 = Tuple.Create(code_module2, vbext_ComponentType.vbext_ct_StdModule);
            var state = Resolve(module1, module2);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable
                && item.Accessibility == Accessibility.Public
                && item.IdentifierName == "Something");

            var usages = declaration.References.Where(item =>
                item.ParentScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
        public void RedimStmt_RedimVariableDeclarationIsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced() As Variant
    ReDim referenced(referenced TO referenced, referenced), referenced(referenced)
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(6, declaration.References.Count());
        }

        [TestMethod]
        public void OpenStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Open referenced For Binary Access Read Lock Read As #referenced Len = referenced
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(3, declaration.References.Count());
        }

        [TestMethod]
        public void CloseStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Close referenced, referenced
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(2, declaration.References.Count());
        }

        [TestMethod]
        public void SeekStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Seek #referenced, referenced
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(2, declaration.References.Count());
        }

        [TestMethod]
        public void LockStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Lock referenced, referenced To referenced
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(3, declaration.References.Count());
        }

        [TestMethod]
        public void UnlockStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Unlock referenced, referenced To referenced
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(3, declaration.References.Count());
        }

        [TestMethod]
        public void LineInputStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Line Input #referenced, referenced
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(2, declaration.References.Count());
        }

        [TestMethod]
        public void WidthStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Width #referenced, referenced
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(2, declaration.References.Count());
        }

        [TestMethod]
        public void PrintStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Print #referenced,,referenced; SPC(referenced), TAB(referenced)
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(4, declaration.References.Count());
        }

        [TestMethod]
        public void WriteStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Write #referenced,,referenced; SPC(referenced), TAB(referenced)
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(4, declaration.References.Count());
        }

        [TestMethod]
        public void InputStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Input #referenced,referenced
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(2, declaration.References.Count());
        }

        [TestMethod]
        public void PutStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Put referenced,referenced,referenced
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(3, declaration.References.Count());
        }

        [TestMethod]
        public void GetStmt_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Get referenced,referenced,referenced
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(3, declaration.References.Count());
        }

        [TestMethod]
        public void CircleSpecialForm_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Me.Circle Step(referenced, referenced), referenced, referenced, referenced, referenced, referenced
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(7, declaration.References.Count());
        }

        [TestMethod]
        public void ScaleSpecialForm_IsReferenceToLocalVariable()
        {
            // arrange
            var code = @"
Public Sub Test()
    Dim referenced As Integer
    Scale (referenced, referenced)-(referenced, referenced)
End Sub
";
            // act
            var state = Resolve(code);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "referenced");

            Assert.AreEqual(4, declaration.References.Count());
        }

        // Ignored because handling forms/hierarchies is an open issue.
        [Ignore]
        [TestMethod]
        public void GivenControlDeclaration_ResolvesUsageInCodeBehind()
        {
            var code = @"
Public Sub DoSomething()
    TextBox1.Height = 20
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", vbext_ProjectProtection.vbext_pp_none);
            var form = project.MockUserFormBuilder("Form1", code).AddControl("TextBox1").Build();
            project.AddComponent(form);
            builder.AddProject(project.Build());
            var vbe = builder.Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState());

            parser.Parse();
            if (parser.State.Status == ParserState.ResolverError)
            {
                Assert.Fail("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }

            var declaration = parser.State.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Control
                && item.IdentifierName == "TextBox1");

            var usages = declaration.References.Where(item =>
                item.ParentNonScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code_class1, code_class2);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Project
                && item.IdentifierName == "TestProject1");

            var usages = declaration.References.Where(item =>
                item.ParentNonScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code_class1, code_class2);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.ClassModule
                && item.IdentifierName == "Class2");

            var usages = declaration.References.Where(item =>
                item.ParentNonScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
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
            var state = Resolve(code_class1, code_class2);

            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.UserDefinedType
                && item.IdentifierName == "TFoo");

            var usages = declaration.References.Where(item =>
                item.ParentNonScoping.IdentifierName == "DoSomething");

            Assert.AreEqual(1, usages.Count());
        }

        [TestMethod]
        public void QualifiedSetStatement_FirstSectionDoesNotHaveAssignmentFlag()
        {
            // arrange
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
            // act
            var state = Resolve(variableDeclarationClass, classVariableDeclarationClass, variableCallClass);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "myClassN");

            Assert.IsFalse(declaration.References.ElementAt(0).IsAssignment);
        }

        [TestMethod]
        public void QualifiedSetStatement_MiddleSectionDoesNotHaveAssignmentFlag()
        {
            // arrange
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
            // act
            var state = Resolve(variableDeclarationClass, classVariableDeclarationClass, variableCallClass);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "myClass");

            Assert.IsFalse(declaration.References.ElementAt(0).IsAssignment);
        }

        [TestMethod]
        public void QualifiedSetStatement_LastSectionHasAssignmentFlag()
        {
            // arrange
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
            // act
            var state = Resolve(variableDeclarationClass, classVariableDeclarationClass, variableCallClass);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

            Assert.IsTrue(declaration.References.ElementAt(0).IsAssignment);
        }

        [TestMethod]
        public void SetStatement_HasAssignmentFlag()
        {
            // arrange
            var variableDeclarationClass = @"
Public foo As Variant

Public Sub bar()
    Set foo = New Class2
End Sub
";
            // act
            var state = Resolve(variableDeclarationClass, string.Empty);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

            Assert.IsTrue(declaration.References.ElementAt(0).IsAssignment);
        }

        [TestMethod]
        public void ImplicitLetStatement_HasAssignmentFlag()
        {
            // arrange
            var variableDeclarationClass = @"
Public foo As Boolean

Public Sub bar()
    foo = True
End Sub
";
            // act
            var state = Resolve(variableDeclarationClass);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

            Assert.IsTrue(declaration.References.ElementAt(0).IsAssignment);
        }

        [TestMethod]
        public void ExplicitLetStatement_HasAssignmentFlag()
        {
            // arrange
            var variableDeclarationClass = @"
Public foo As Boolean

Public Sub bar()
    Let foo = True
End Sub
";
            // act
            var state = Resolve(variableDeclarationClass);

            // assert
            var declaration = state.AllUserDeclarations.Single(item =>
                item.DeclarationType == DeclarationType.Variable && item.IdentifierName == "foo");

            Assert.IsTrue(declaration.References.ElementAt(0).IsAssignment);
        }
    }
}

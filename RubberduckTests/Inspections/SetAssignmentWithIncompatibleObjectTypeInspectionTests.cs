using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Moq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.TypeResolvers;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class SetAssignmentWithIncompatibleObjectTypeInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void AssignmentToNotImplementedInterface_Result()
        {
            const string interface1 =
                @"Public Sub DoIt()
End Sub";

            const string class1 =
                @"
Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim intrfc As Interface1
    Dim cls As Class1

    Set cls = new Class1
    Set intrfc = cls
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Interface1", interface1, ComponentType.ClassModule),
                ("Class1", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule));

            Assert.AreEqual(1,inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentToInterfaceIncompatibleWithDeclaredTypeButNotWithUnderlyingType_Result()
        {
            const string interface1 =
                @"Public Sub DoIt()
End Sub";

            const string interface2 =
                @"Public Sub DoSomething()
End Sub";

            const string class1 =
                @"Implements Interface1
Implements Interface2

Private Sub Interface1_DoIt()
End Sub

Private Sub Interface2_DoSomething()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim intrfc As Interface1
    Dim otherIntrfc As Interface2

    Set otherIntrfc = new Class1
    Set intrfc = otherIntrfc
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Interface1", interface1, ComponentType.ClassModule),
                ("Interface2", interface2, ComponentType.ClassModule),
                ("Class1", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule));

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentToImplementedInterface_NoResult()
        {
            const string interface1 =
                @"Public Sub DoIt()
End Sub";
            
                const string class1 =
                @"Implements Interface1

Private Sub Interface1_DoIt()
End Sub
";

                const string consumerModule =
                    @"
Private Sub TestIt()
    Dim intrfc As Interface1
    Dim cls As Class1

    Set cls = new Class1
    Set intrfc = cls
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                    ("Interface1", interface1, ComponentType.ClassModule),
                    ("Class1", class1, ComponentType.ClassModule),
                    ("Module1", consumerModule, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentToSameClass_NoResult()
        {
            const string class1 =
                @"
Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim cls As Class1
    Dim otherCls As Class1

    Set otherCls = new Class1
    Set cls = otherCls
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentToSameClass_InconsistentlyQualified_NoResult()
        {
            const string class1 =
                @"
Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim cls As TestProject1.Class1
    Dim otherCls As Class1

    Set otherCls = new Class1
    Set cls = otherCls
    Set otherCls = cls 
End Sub
";
            var modules = new(string, string, ComponentType)[] 
            {
                ("Class1", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule),
            };

            Assert.AreEqual(0, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentToOtherClassWithSameName_OneResultEach()
        {
            const string class1 =
                @"Attribute VB_Exposed = True
Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim cls As Project2.Class1
    Dim otherCls As Class1

    Set otherCls = new Class1
    Set cls = otherCls
    Set otherCls = cls 
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("Project2", "project2path", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddProjectToVbeBuilder()
                .ProjectBuilder("Project1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, consumerModule)
                .AddReference("Project2", "project2path", 0, 0, false, ReferenceKind.Project)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var inspectionResults = InspectionResults(vbe).ToList();

            Assert.AreEqual(2,inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void LegalDowncastFromImplementedInterface_NoResult()
        {
            const string interface1 =
                @"Public Sub DoIt()
End Sub";

            const string class1 =
                @"Implements Interface1

Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim intrfc As Interface1
    Dim cls As Class1

    Set intrfc = new Class1
    Set cls = intrfc
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Interface1", interface1, ComponentType.ClassModule),
                ("Class1", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        //We cannot know whether a downcast is legal at compile time.
        public void IllegalDowncastFromImplementedInterface_NoResult()
        {
            const string interface1 =
                @"Public Sub DoIt()
End Sub";

            const string class1 =
                @"Implements Interface1

Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim intrfc As Interface1
    Dim cls As Class1

    Set intrfc = new Class2
    Set cls = intrfc
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Interface1", interface1, ComponentType.ClassModule),
                ("Class1", class1, ComponentType.ClassModule),
                ("Class2", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentToObject_NoResult()
        {
            const string class1 =
                @"
Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim cls As Object
    Dim otherCls As Class1

    Set otherCls = new Class1
    Set cls = otherCls
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentToVariant_NoResult()
        {
            const string class1 =
                @"
Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim cls As Variant
    Dim otherCls As Class1

    Set otherCls = new Class1
    Set cls = otherCls
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentToIUnknown_NoResult()
        {
            const string class1 =
                @"
Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim cls As IUnknown
    Dim otherCls As Class1

    Set otherCls = new Class1
    Set cls = otherCls
End Sub
";
            var testVbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, consumerModule)
                .AddReference(ReferenceLibrary.StdOle)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var inspectionResults = InspectionResults(testVbe);

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentOfObject_NoResult()
        {
            const string class1 =
                @"
Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim cls As Class1
    Dim otherCls As Object

    Set otherCls = new Class2
    Set cls = otherCls
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1, ComponentType.ClassModule),
                ("Class2", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentOfVariant_NoResult()
        {
            const string class1 =
                @"
Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim cls As Class1
    Dim otherCls As Variant

    Set otherCls = new Class2
    Set cls = otherCls
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1, ComponentType.ClassModule),
                ("Class2", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentOfIUnknown_NoResult()
        {
            const string class1 =
                @"
Private Sub Interface1_DoIt()
End Sub
";

            const string consumerModule =
                @"
Private Sub TestIt()
    Dim cls As Class1
    Dim otherCls As IUnknown

    Set otherCls = new Class2
    Set cls = otherCls
End Sub
";
            var testVbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Class2", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, consumerModule)
                .AddReference(ReferenceLibrary.StdOle)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var inspectionResults = InspectionResults(testVbe);

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentOfMeToProperlyTypesVariable_NoResult()
        {
            const string interface1 =
                @"
Private Sub DoIt()
End Sub
";
            const string class1 =
                @"Implements Interface1
Private Sub Interface1_DoIt()
End Sub

Public Sub AssignIt()
    Dim cls As Interface1
    Set cls = Me
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1, ComponentType.ClassModule),
                ("Interface1", interface1, ComponentType.ClassModule));

            Assert.IsFalse(inspectionResults.Any());
        }

        [Test]
        [Category("Inspections")]
        public void AssignmentOfMeToImproperlyTypesVariable_Result()
        {
            const string interface1 =
                @"
Private Sub DoIt()
End Sub
";
            const string class1 =
                @"
Private Sub Interface1_DoIt()
End Sub

Public Sub AssignIt()
    Dim cls As Interface1
    Set cls = Me
End Sub
";

            var inspectionResults = InspectionResultsForModules(
                ("Class1", class1, ComponentType.ClassModule),
                ("Interface1", interface1, ComponentType.ClassModule));

            Assert.AreEqual(1,inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Class1", "TestProject1.Class1", 0)]
        [TestCase("Interface1", "TestProject1.Class1", 0)]
        [TestCase("Class1", "TestProject1.Interface1", 0)]
        [TestCase("Variant", "Class1", 0)] //Tokens.Variant cannot be used here because it is not a constant expression.
        [TestCase("Object", "Class1", 0)]
        [TestCase("Class1", "Variant", 0)]
        [TestCase("Class1", "Object", 0)]
        [TestCase("Class1", "TestProject1.SomethingIncompatible", 1)]
        [TestCase("Class1", "SomethingDifferent", 1)]
        [TestCase("TestProject1.Class1", "OtherProject.Class1", 1)]
        [TestCase("TestProject1.Interface1", "OtherProject.Class1", 1)]
        [TestCase("TestProject1.Class1", "OtherProject.Interface1", 1)]
        [TestCase("Class1", "OtherProject.Class1", 1)]
        [TestCase("Interface1", "OtherProject.Class1", 1)]
        [TestCase("Class1", "OtherProject.Interface1", 1)]
        [TestCase("Class1", SetTypeResolver.NotAnObject, 1)] //The RHS is not even an object. (Will show as type NotAnObject in the result.) 
        [TestCase("Class1", null, 0)] //We could not resolve the Set type, so we do not return a result. 
        public void MockedSetTypeEvaluatorTest_Function(string lhsTypeName, string expressionFullTypeName, int expectedResultsCount)
        {
            const string interface1 =
                @"
Private Sub Foo() 
End Sub
";
            const string class1 =
                @"Implements Interface1

Private Sub Interface1_Foo()
End Sub
";

            var module1 =
                $@"
Private Function Cls() As {lhsTypeName}
    Set Cls = expression
End Function
";
            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", class1, ComponentType.ClassModule),
                ("Interface1", interface1, ComponentType.ClassModule),
                ("Module1", module1, ComponentType.StandardModule),
            };

            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;

            var setTypeResolverMock = new Mock<ISetTypeResolver>();
            setTypeResolverMock.Setup(m =>
                    m.SetTypeName(It.IsAny<VBAParser.ExpressionContext>(), It.IsAny<QualifiedModuleName>()))
                .Returns((VBAParser.ExpressionContext context, QualifiedModuleName qmn) => expressionFullTypeName);

            var inspectionResults = InspectionResults(vbe, setTypeResolverMock.Object).ToList();

            Assert.AreEqual(expectedResultsCount, inspectionResults.Count);
        }

        [Test]
        [Category("Inspections")]
        [TestCase("IUnknown", "Class1", 0)]
        [TestCase("Class1", ":stdole.IUnknown", 0)]
        public void MockedSetTypeEvaluatorTest_Function_IUnknown(string lhsTypeName, string expressionFullTypeName, int expectedResultsCount)
        {
            const string interface1 =
                @"
Private Sub Foo() 
End Sub
";
            const string class1 =
                @"Implements Interface1

Private Sub Interface1_Foo()
End Sub
";

            var module1 =
                $@"
Private Function Cls() As {lhsTypeName}
    Set Cls = expression
End Function
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Interface1", ComponentType.ClassModule, interface1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddReference(ReferenceLibrary.StdOle)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeResolverMock = new Mock<ISetTypeResolver>();
            setTypeResolverMock.Setup(m =>
                    m.SetTypeName(It.IsAny<VBAParser.ExpressionContext>(), It.IsAny<QualifiedModuleName>()))
                .Returns((VBAParser.ExpressionContext context, QualifiedModuleName qmn) => expressionFullTypeName);

            var inspectionResults = InspectionResults(vbe, setTypeResolverMock.Object).ToList();

            Assert.AreEqual(expectedResultsCount, inspectionResults.Count);
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Class1", "TestProject1.Class1", 0)]
        [TestCase("Interface1", "TestProject1.Class1", 0)]
        [TestCase("Class1", "TestProject1.Interface1", 0)]
        [TestCase("Variant", "Class1", 0)] //Tokens.Variant cannot be used here because it is not a constant expression.
        [TestCase("Object", "Class1", 0)]
        [TestCase("Class1", "Variant", 0)]
        [TestCase("Class1", "Object", 0)]
        [TestCase("Class1", "TestProject1.SomethingIncompatible", 1)]
        [TestCase("Class1", "SomethingDifferent", 1)]
        [TestCase("TestProject1.Class1", "OtherProject.Class1", 1)]
        [TestCase("TestProject1.Interface1", "OtherProject.Class1", 1)]
        [TestCase("TestProject1.Class1", "OtherProject.Interface1", 1)]
        [TestCase("Class1", "OtherProject.Class1", 1)]
        [TestCase("Interface1", "OtherProject.Class1", 1)]
        [TestCase("Class1", "OtherProject.Interface1", 1)]
        [TestCase("Class1", SetTypeResolver.NotAnObject, 1)] //The RHS is not even an object. (Will show as type NotAnObject in the result.) 
        [TestCase("Class1", null, 0)] //We could not resolve the Set type, so we do not return a result. 
        public void MockedSetTypeEvaluatorTest_PropertyGet(string lhsTypeName, string expressionFullTypeName, int expectedResultsCount)
        {
            const string interface1 =
                @"
Private Sub Foo() 
End Sub
";
            const string class1 =
                @"Implements Interface1

Private Sub Interface1_Foo()
End Sub
";

            var module1 =
                $@"
Private Property Get Cls() As {lhsTypeName}
    Set Cls = expression
End Property
";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", class1, ComponentType.ClassModule),
                ("Interface1", interface1, ComponentType.ClassModule),
                ("Module1", module1, ComponentType.StandardModule),
            };

            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;

            var setTypeResolverMock = new Mock<ISetTypeResolver>();
            setTypeResolverMock.Setup(m =>
                    m.SetTypeName(It.IsAny<VBAParser.ExpressionContext>(), It.IsAny<QualifiedModuleName>()))
                .Returns((VBAParser.ExpressionContext context, QualifiedModuleName qmn) => expressionFullTypeName);

            var inspectionResults = InspectionResults(vbe, setTypeResolverMock.Object).ToList();

            Assert.AreEqual(expectedResultsCount, inspectionResults.Count);
        }

        [Test]
        [Category("Inspections")]
        [TestCase("IUnknown", "Class1", 0)]
        [TestCase("Class1", ":stdole.IUnknown", 0)]
        public void MockedSetTypeEvaluatorTest_PropertyGet_IUnknown(string lhsTypeName, string expressionFullTypeName, int expectedResultsCount)
        {
            const string interface1 =
                @"
Private Sub Foo() 
End Sub
";
            const string class1 =
                @"Implements Interface1

Private Sub Interface1_Foo()
End Sub
";

            var module1 =
                $@"
Private Property Get Cls() As {lhsTypeName}
    Set Cls = expression
End Property
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Interface1", ComponentType.ClassModule, interface1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddReference(ReferenceLibrary.StdOle)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeResolverMock = new Mock<ISetTypeResolver>();
            setTypeResolverMock.Setup(m =>
                    m.SetTypeName(It.IsAny<VBAParser.ExpressionContext>(), It.IsAny<QualifiedModuleName>()))
                .Returns((VBAParser.ExpressionContext context, QualifiedModuleName qmn) => expressionFullTypeName);

            var inspectionResults = InspectionResults(vbe, setTypeResolverMock.Object).ToList();

            Assert.AreEqual(expectedResultsCount, inspectionResults.Count);
        }

        [Test]
        [Category("Inspections")]
        [TestCase("Class1", "TestProject1.Class1", 0)]
        [TestCase("Interface1", "TestProject1.Class1", 0)]
        [TestCase("Class1", "TestProject1.Interface1", 0)]
        [TestCase("Variant", "Class1", 0)] //Tokens.Variant cannot be used here because it is not a constant expression.
        [TestCase("Object", "Class1", 0)]
        [TestCase("Class1", "Variant", 0)]
        [TestCase("Class1", "Object", 0)]
        [TestCase("Class1", "TestProject1.SomethingIncompatible", 1)]
        [TestCase("Class1", "SomethingDifferent", 1)]
        [TestCase("TestProject1.Class1", "OtherProject.Class1", 1)]
        [TestCase("TestProject1.Interface1", "OtherProject.Class1", 1)]
        [TestCase("TestProject1.Class1", "OtherProject.Interface1", 1)]
        [TestCase("Class1", "OtherProject.Class1", 1)]
        [TestCase("Interface1", "OtherProject.Class1", 1)]
        [TestCase("Class1", "OtherProject.Interface1", 1)]
        [TestCase("Class1", SetTypeResolver.NotAnObject, 1)] //The RHS is not even an object. (Will show as type NotAnObject in the result.) 
        [TestCase("Class1", null, 0)] //We could not resolve the Set type, so we do not return a result. 
        public void MockedSetTypeEvaluatorTest_Variable(string lhsTypeName, string expressionFullTypeName, int expectedResultsCount)
        {
            const string interface1 =
                @"
Private Sub Foo() 
End Sub
";
            const string class1 =
                @"Implements Interface1

Private Sub Interface1_Foo()
End Sub
";

            var module1 =
                $@"
Private Sub TestIt()
    Dim cls As {lhsTypeName}

    Set cls = expression
End Sub
";

            var modules = new(string, string, ComponentType)[]
            {
                ("Class1", class1, ComponentType.ClassModule),
                ("Interface1", interface1, ComponentType.ClassModule),
                ("Module1", module1, ComponentType.StandardModule),
            };

            var vbe = MockVbeBuilder.BuildFromModules(modules).Object;

            var setTypeResolverMock = new Mock<ISetTypeResolver>();
            setTypeResolverMock.Setup(m =>
                    m.SetTypeName(It.IsAny<VBAParser.ExpressionContext>(), It.IsAny<QualifiedModuleName>()))
                .Returns((VBAParser.ExpressionContext context, QualifiedModuleName qmn) => expressionFullTypeName);

            var inspectionResults = InspectionResults(vbe, setTypeResolverMock.Object).ToList();

            Assert.AreEqual(expectedResultsCount, inspectionResults.Count);
        }

        [Test]
        [Category("Inspections")]
        [TestCase("IUnknown", "Class1", 0)]
        [TestCase("Class1", ":stdole.IUnknown", 0)]
        public void MockedSetTypeEvaluatorTest_Variable_IUnknown(string lhsTypeName, string expressionFullTypeName, int expectedResultsCount)
        {
            const string interface1 =
                @"
Private Sub Foo() 
End Sub
";
            const string class1 =
                @"Implements Interface1

Private Sub Interface1_Foo()
End Sub
";

            var module1 =
                $@"
Private Sub TestIt()
    Dim cls As {lhsTypeName}

    Set cls = expression
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Interface1", ComponentType.ClassModule, interface1)
                .AddComponent("Module1", ComponentType.StandardModule, module1)
                .AddReference(ReferenceLibrary.StdOle)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var setTypeResolverMock = new Mock<ISetTypeResolver>();
            setTypeResolverMock.Setup(m =>
                    m.SetTypeName(It.IsAny<VBAParser.ExpressionContext>(), It.IsAny<QualifiedModuleName>()))
                .Returns((VBAParser.ExpressionContext context, QualifiedModuleName qmn) => expressionFullTypeName);

            var inspectionResults = InspectionResults(vbe, setTypeResolverMock.Object).ToList();

            Assert.AreEqual(expectedResultsCount, inspectionResults.Count);
        }

        private static IEnumerable<IInspectionResult> InspectionResults(IVBE vbe, ISetTypeResolver setTypeResolver)
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var inspection = InspectionUnderTest(state, setTypeResolver);
                return inspection.GetInspectionResults(CancellationToken.None);
            }
        }

        private static IInspection InspectionUnderTest(RubberduckParserState state, ISetTypeResolver setTypeResolver)
        {
            return new SetAssignmentWithIncompatibleObjectTypeInspection(state, setTypeResolver);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new SetAssignmentWithIncompatibleObjectTypeInspection(state, new SetTypeResolver(state));
        }
    }
}
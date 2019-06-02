using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class SetAssignmentWithIncompatibleObjectTypeInspectionTests
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

            var inspectionResults = InspectionResults(
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

            var inspectionResults = InspectionResults(
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

            var inspectionResults = InspectionResults(
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

            var inspectionResults = InspectionResults(
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
    Dim cls As Project1.Class1
    Dim otherCls As Class1

    Set otherCls = new Class1
    Set cls = otherCls
    Set otherCls = cls 
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("Project1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, class1)
                .AddComponent("Module1", ComponentType.StandardModule, consumerModule)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;

            var inspectionResults = InspectionResults(vbe);

            Assert.IsFalse(inspectionResults.Any());
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

            var inspectionResults = InspectionResults(
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

            var inspectionResults = InspectionResults(
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

            var inspectionResults = InspectionResults(
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

            var inspectionResults = InspectionResults(
                ("Class1", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule));

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

            var inspectionResults = InspectionResults(
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

            var inspectionResults = InspectionResults(
                ("Class1", class1, ComponentType.ClassModule),
                ("Class2", class1, ComponentType.ClassModule),
                ("Module1", consumerModule, ComponentType.StandardModule));

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

            var inspectionResults = InspectionResults(
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

            var inspectionResults = InspectionResults(
                ("Class1", class1, ComponentType.ClassModule),
                ("Interface1", interface1, ComponentType.ClassModule));

            Assert.AreEqual(1,inspectionResults.Count());
        }

        private static IEnumerable<IInspectionResult> InspectionResults(params (string moduleName, string content, ComponentType componentType)[] testModules)
        {
            var vbe = MockVbeBuilder.BuildFromModules(testModules).Object;
            return InspectionResults(vbe);
        }

        private static IEnumerable<IInspectionResult> InspectionResults(IVBE vbe)
        {
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var inspection = InspectionUnderTest(state);
                return inspection.GetInspectionResults(CancellationToken.None);
            }
        }

        private static IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new SetAssignmentWithIncompatibleObjectTypeInspection(state);
        }
    }
}
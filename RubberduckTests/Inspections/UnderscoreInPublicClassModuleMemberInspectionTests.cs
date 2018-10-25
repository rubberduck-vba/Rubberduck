using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Linq;
using System.Threading;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UnderscoreInPublicClassModuleMemberInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void BasicExample_Sub()
        {
            const string inputCode =
                @"Public Sub Test_This_Out()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnderscoreInPublicClassModuleMemberInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void BasicExample_Function()
        {
            const string inputCode =
                @"Public Function Test_This_Out() As Integer
End Function";
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnderscoreInPublicClassModuleMemberInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void BasicExample_Property()
        {
            const string inputCode =
                @"Public Property Get Test_This_Out() As Integer
End Property";
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnderscoreInPublicClassModuleMemberInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void StandardModule()
        {
            const string inputCode =
                @"Public Sub Test_This_Out()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnderscoreInPublicClassModuleMemberInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void NoUnderscore()
        {
            const string inputCode =
                @"Public Sub Foo()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnderscoreInPublicClassModuleMemberInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void FriendMember_WithUnderscore()
        {
            const string inputCode =
                @"Friend Sub Test_This_Out()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnderscoreInPublicClassModuleMemberInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void Implicit_WithUnderscore()
        {
            const string inputCode =
                @"Sub Test_This_Out()
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, ComponentType.ClassModule, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnderscoreInPublicClassModuleMemberInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ImplementsInterface()
        {
            const string inputCode1 =
       @"Public Sub Foo()
End Sub";

            //Expectation
            const string inputCode2 =
                @"Implements Class1

Public Sub Class1_Foo()
    Err.Raise 5 'TODO implement interface member
End Sub
";
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode1)
                .AddComponent("Class2", ComponentType.ClassModule, inputCode2)
                .Build();
            var vbe = builder.AddProject(project).Build();
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UnderscoreInPublicClassModuleMemberInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }
    }
}

using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class ImplicitDefaultMemberAssignmentInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void ImplicitDefaultMemberAssignment_ReturnsResult()
        {
            const string defaultMemberClassCode = @"
Public Property Let Foo(bar As Long)
Attribute Foo.VB_UserMemId = 0
End Property
";

            const string inputCode = @"
Public Sub Foo()
    Dim bar As Class1
    bar = 42
End Sub
";

            var inspectionResults = GetInspectionResults(defaultMemberClassCode, inputCode);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitDefaultMemberAssignmentOnObject_ReturnsResult()
        {
            const string defaultMemberClassCode = @"
Public Property Let Foo(bar As Long)
Attribute Foo.VB_UserMemId = 0
End Property
";

            const string inputCode = @"
Public Sub Foo()
    Dim bar As Object
    bar = 42
End Sub
";

            var inspectionResults = GetInspectionResults(defaultMemberClassCode, inputCode);

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitDefaultMemberAssignment_IgnoredDoesNotReturnResult()
        {
            const string defaultMemberClassCode = @"
Public Property Let Foo(bar As Long)
Attribute Foo.VB_UserMemId = 0
End Property
";

            const string inputCode = @"
Public Sub Foo(bar As Class1)
    '@Ignore ImplicitDefaultMemberAssignment
    bar = 42
End Sub
";

            var inspectionResults = GetInspectionResults(defaultMemberClassCode, inputCode);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitDefaultMemberAssignment_ExplicitCallDoesNotReturnResult()
        {
            const string defaultMemberClassCode = @"
Public Property Let Foo(bar As Long)
Attribute Foo.VB_UserMemId = 0
End Property
";

            const string inputCode = @"
Public Sub Foo(bar As Class1)
    bar.Foo = 42
End Sub
";

            var inspectionResults = GetInspectionResults(defaultMemberClassCode, inputCode);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void ImplicitDefaultMemberAssignment_ExplicitLetDoesNotReturnResult()
        {
            const string defaultMemberClassCode = @"
Public Property Let Foo(bar As Long)
Attribute Foo.VB_UserMemId = 0
End Property
";

            const string inputCode = @"
Public Sub Foo(bar As Class1)
    Let bar = 42
End Sub
";

            var inspectionResults = GetInspectionResults(defaultMemberClassCode, inputCode);

            Assert.AreEqual(0, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "ImplicitDefaultMemberAssignmentInspection";
            var inspection = new ImplicitDefaultMemberAssignmentInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        private IEnumerable<IInspectionResult> GetInspectionResults(string defaultMemberClassCode, string moduleCode)
        {
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, defaultMemberClassCode)
                .AddComponent("Module1", ComponentType.StandardModule, moduleCode)
                .AddProjectToVbeBuilder()
                .Build()
                .Object;
            using (var state = MockParser.CreateAndParse(vbe))
            {
                var inspection = InspectionUnderTest(state);
                return inspection.GetInspectionResults(CancellationToken.None);
            }
        }

        private IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new ImplicitDefaultMemberAssignmentInspection(state);
        }
    }
}

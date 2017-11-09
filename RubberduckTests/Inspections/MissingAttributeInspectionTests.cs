using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class MissingAttributeInspectionTests
    {
        [TestCategory("Inspections")]
        [TestMethod]
        public void NoResultGivenNoAttribute_NoAnnotation()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Attribute VB_GlobalNameSpace = False              ' (ignored)
Attribute VB_Creatable = False                    ' (ignored)
Attribute VB_PredeclaredId = False                ' Must be False
Attribute VB_Exposed = False                      ' Must be False
Option Explicit
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAttributeInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestCategory("Inspections")]
        [TestMethod]
        public void NoResultGivenNonAttributeAnnotation()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Attribute VB_GlobalNameSpace = False              ' (ignored)
Attribute VB_Creatable = False                    ' (ignored)
Attribute VB_PredeclaredId = False                ' Must be False
Attribute VB_Exposed = False                      ' Must be False
Option Explicit

'@TestMethod
Public Sub Test()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAttributeInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestCategory("Inspections")]
        [TestMethod]
        public void HasResultGivenPredeclaredIdAnnotation_WithoutAttribute()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Attribute VB_GlobalNameSpace = False              ' (ignored)
Attribute VB_Creatable = False                    ' (ignored)
Attribute VB_PredeclaredId = False                ' Must be False
Attribute VB_Exposed = False                      ' Must be False
Option Explicit
'@PredeclaredId
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAttributeInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestCategory("Inspections")]
        [TestMethod]
        public void NoResultGivenPredeclaredIdAttribute_WithAnnotation()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Attribute VB_GlobalNameSpace = False              ' (ignored)
Attribute VB_Creatable = False                    ' (ignored)
Attribute VB_PredeclaredId = True                 ' Must be True
Attribute VB_Exposed = False                      ' Must be False
Option Explicit
'@PredeclaredId
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAttributeInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestCategory("Inspections")]
        [TestMethod]
        public void NoResultGivenExposedAttribute_NoAnnotation()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Attribute VB_GlobalNameSpace = False              ' (ignored)
Attribute VB_Creatable = False                    ' (ignored)
Attribute VB_PredeclaredId = False                ' Must be False
Attribute VB_Exposed = True                       ' Must be True
Option Explicit
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAttributeInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestCategory("Inspections")]
        [TestMethod]
        public void HasResultGivenExposedAnnotation_WithoutAttribute()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Attribute VB_GlobalNameSpace = False              ' (ignored)
Attribute VB_Creatable = False                    ' (ignored)
Attribute VB_PredeclaredId = False                ' Must be False
Attribute VB_Exposed = False                      ' Must be False
Option Explicit
'@Exposed
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAttributeInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [TestCategory("Inspections")]
        [TestMethod]
        public void NoResultGivenMemberDescriptionAttribute_NoAnnotation()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Attribute VB_GlobalNameSpace = False              ' (ignored)
Attribute VB_Creatable = False                    ' (ignored)
Attribute VB_PredeclaredId = False                ' Must be False
Attribute VB_Exposed = False                      ' Must be False
Option Explicit

Sub DoSomething()
Attribute DoSomething.VB_Description = ""Does something""
End Sub
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAttributeInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(inspectionResults.Any());
            }
        }

        [TestCategory("Inspections")]
        [TestMethod]
        public void HasResultGivenDescriptionAnnotation_WithoutAttribute()
        {
            const string testModuleName = "Test";
            const string inputCode = @"
VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = """ + testModuleName + @"""   ' (ignored)
Attribute VB_GlobalNameSpace = False              ' (ignored)
Attribute VB_Creatable = False                    ' (ignored)
Attribute VB_PredeclaredId = False                ' Must be False
Attribute VB_Exposed = False                      ' Must be False
Option Explicit

'@Description(""Does something"")
Sub DoSomething()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromSingleModule(inputCode, testModuleName, ComponentType.ClassModule, out _);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new MissingAttributeInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }
    }
}
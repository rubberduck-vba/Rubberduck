﻿using System.Linq;
using System.Threading;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RubberduckTests.Mocks;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Parsing.Inspections.Resources;

namespace RubberduckTests.Inspections
{
    [TestClass, Ignore]
    public class EmptyElseBlockInspectionTests
    {
        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionType()
        {
            var inspection = new EmptyElseBlockInspection(null);
            var expectedInspection = CodeInspectionType.MaintainabilityAndReadabilityIssues;

            Assert.AreEqual(expectedInspection, inspection.InspectionType);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void InspectionName()
        {
            const string expectedName = nameof(EmptyElseBlockInspection);
            var inspection = new EmptyElseBlockInspection(null);

            Assert.AreEqual(expectedName, inspection.Name);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_DoesntFireOnEmptyIfBlock()
        {
            const string inputcode =
                @"Sub Foo()
    If True Then
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyElseBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                const int expectedCount = 0;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasNoContent()
        {
            const string inputcode =
                @"Sub Foo()
    If True Then
    Else
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyElseBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                const int expectedCount = 1;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasQuoteComment()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        'Some Comment
    End If
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyElseBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                const int expectedCount = 1;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasRemComment()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Rem Some Comment
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyElseBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                const int expectedCount = 1;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasVariableDeclaration()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Dim d
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyElseBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                const int expectedCount = 1;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasConstDeclaration()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Const c = 0
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyElseBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                const int expectedCount = 1;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasWhitespace()
        {
            const string inputcode =
                @"Sub Foo()
    If True Then
    Else
    
    
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyElseBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                const int expectedCount = 1;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasDeclarationStatement()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Dim d
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyElseBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                const int expectedCount = 1;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void EmptyElseBlock_HasExecutableStatement()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
        Dim d
        d = 0
    End If
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new EmptyElseBlockInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var actualResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;
                const int expectedCount = 0;

                Assert.AreEqual(expectedCount, actualResults.Count());
            }
        }
    }
}

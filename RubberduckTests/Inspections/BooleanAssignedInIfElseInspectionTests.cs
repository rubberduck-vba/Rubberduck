using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class BooleanAssignedInIfElseInspectionTests
    {
        [Test]
        [Category("Inspections")]
        public void Simple()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d As Boolean
    If True Then
        d = True
    Else
        d = False
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void QualifiedName()
        {
            const string inputcode =
                @"Sub Foo()
    If True Then
        Fizz.Buzz = True
    Else
        Fizz.Buzz = False
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void MultipleResults()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d As Boolean
    If True Then
        d = True
    Else
        d = False
    EndIf

    If True Then
        d = False
    Else
        d = True
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(2, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void AssignsInteger()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = 0
    Else
        d = 1
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(results.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void AssignsTheSameValue()       // worthy of an inspection in its own right
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = True
    Else
        d = True
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(results.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void AssignsToDifferentVariables()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d1, d2
    If True Then
        d1 = True
    Else
        d2 = False
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(results.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ConditionalContainsElseIfBlock()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = True
    ElseIf False Then
        d = True
    Else
        d = False
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(results.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void ConditionalDoesNotContainElseBlock()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = True
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(results.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void IsIgnored()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    '@Ignore BooleanAssignedInIfElse
    If True Then
        d = True
    Else
        d = False
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.IsFalse(results.Any());
            }
        }

        [Test]
        [Category("Inspections")]
        public void BlockContainsPrefixComment()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        ' test
        d = True
    Else
        d = False
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void BlockContainsPostfixComment()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = True
        ' test
    Else
        d = False
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, results.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void BlockContainsEOLComment()
        {
            const string inputcode =
                @"Sub Foo()
    Dim d
    If True Then
        d = True    ' test
    Else
        d = False
    EndIf
End Sub";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputcode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var results = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, results.Count());
            }
        }
    }
}
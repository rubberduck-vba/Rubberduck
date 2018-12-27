using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using RubberduckTests.Mocks;
using System.Threading;
using System.Linq;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class DefTypeStatementInspectionTests
    {
        [Test]
        [TestCase("Bool")]
        [TestCase("Byte")]
        [TestCase("Int")]
        [TestCase("Lng")]
        [TestCase("Cur")]
        [TestCase("Sng")]
        [TestCase("Dbl")]
        [TestCase("Date")]
        [TestCase("Str")]
        [TestCase("Obj")]
        [TestCase("Var")]
        [Category("Inspections")]
        public void DefType_SingleResultFound(string type)
        {
            const string inputCode =
@"Def{0} A
Public Function aFoo()
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(inputCode, type), out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DefTypeStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(1, inspectionResults.Count());
            }
        }

        [Test]
        [TestCase("Bool")]
        [TestCase("Byte")]
        [TestCase("Int")]
        [TestCase("Lng")]
        [TestCase("Cur")]
        [TestCase("Sng")]
        [TestCase("Dbl")]
        [TestCase("Date")]
        [TestCase("Str")]
        [TestCase("Obj")]
        [TestCase("Var")]
        [Category("Inspections")]
        public void DefType_SingleResultIgnored(string type)
        {
            const string inputCode =
@"'@Ignore DefTypeStatement
Def{0} F
Public Function FunctionWontBeFoundInResult()
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(string.Format(inputCode, type), out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DefTypeStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void DefType_AllDefTypeAreFound()
        {
            const string inputCode =
@"DefBool A
DefByte B
DefInt C
DefLng D
DefCur E
DefSng F
DefDbl G
DefDate H
DefStr I
DefObj J
DefVar K
Public Function Zoo()
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DefTypeStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(11, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void DefType_AllDefTypeAreIgnored()
        {
            const string inputCode =
@"'@IgnoreModule DefTypeStatement
DefBool A
DefByte B
DefInt C
DefLng D
DefCur E
DefSng F
DefDbl G
DefDate H
DefStr I
DefObj J
DefVar K
Public Function Zoo()
End Function";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out _);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new DefTypeStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                Assert.AreEqual(0, inspectionResults.Count());
            }
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "DefTypeStatementInspection";
            var inspection = new DefTypeStatementInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }
    }
}
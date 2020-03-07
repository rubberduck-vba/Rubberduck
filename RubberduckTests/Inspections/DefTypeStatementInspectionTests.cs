using NUnit.Framework;
using System.Linq;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class DefTypeStatementInspectionTests : InspectionTestsBase
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

            Assert.AreEqual(1, InspectionResultsForStandardModule(string.Format(inputCode, type)).Count());
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

            Assert.AreEqual(0, InspectionResultsForStandardModule(string.Format(inputCode, type)).Count());
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

            Assert.AreEqual(11, InspectionResultsForStandardModule(inputCode).Count());
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

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            const string inspectionName = "DefTypeStatementInspection";
            var inspection = new DefTypeStatementInspection(null);

            Assert.AreEqual(inspectionName, inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new DefTypeStatementInspection(state);
        }
    }
}
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.Parsing.VBA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class MisleadingByRefParameterInspectionTests : InspectionTestsBase
    {
        [TestCase("Property Let Fizz(ByRef arg1 As Integer)\r\nEnd Property", 1)]
        [TestCase("Property Let Fizz(arg1 As Integer)\r\nEnd Property", 0)]
        [TestCase("Property Let Fizz(ByVal arg1 As Integer)\r\nEnd Property", 0)]
        [TestCase("Property Set Fizz(ByRef arg1 As Object)\r\nEnd Property", 1)]
        [TestCase("Property Set Fizz(arg1 As Object)\r\nEnd Property", 0)]
        [TestCase("Property Set Fizz(ByVal arg1 As Object)\r\nEnd Property", 0)]
        [Category("QuickFixes")]
        [Category(nameof(MisleadingByRefParameterInspection))]
        public void AllParamMechanisms(string inputCode, int expectedCount)
        {
            Assert.AreEqual(expectedCount, InspectionResultsForStandardModule(inputCode).Count());
        }

        [TestCase("arg")]
        [TestCase("ByRef arg")]
        [Category("QuickFixes")]
        [Category(nameof(MisleadingByRefParameterInspection))]
        public void UserDefinedTypeEdgeCase(string parameterMechanismAndParam)
        {
            var inputCode =
$@"
Option Explicit

Public Type TestType
    FirstValue As Long
End Type

Private this As TestType

Public Property Get UserDefinedType() As TestType
    UserDefinedType = this
End Property

Public Property Let UserDefinedType({parameterMechanismAndParam} As TestType)
    this = arg
End Property
";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5628
        [TestCase("ArrayToStore")]
        [TestCase("ByRef ArrayToStore")]
        [Category("QuickFixes")]
        [Category(nameof(MisleadingByRefParameterInspection))]
        public void ArrayEdgeCase(string parameterMechanismAndParam)
        {
            var inputCode =
$@"
Option Explicit

Private InternalArray() As Variant

Public Property Get StoredArray() As Variant()
    StoredArray = InternalArray
End Property

Public Property Let StoredArray({parameterMechanismAndParam}() As Variant)
    InternalArray = ArrayToStore
End Property
";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("QuickFixes")]
        [Category(nameof(MisleadingByRefParameterInspection))]
        public void InspectionName()
        {
            var inspection = new MisleadingByRefParameterInspection(null);

            Assert.AreEqual(nameof(MisleadingByRefParameterInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new MisleadingByRefParameterInspection(state);
        }
    }
}

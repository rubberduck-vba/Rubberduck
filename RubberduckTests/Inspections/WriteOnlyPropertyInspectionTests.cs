using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class WriteOnlyPropertyInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void WriteOnlyProperty_ReturnsResult_Let()
        {
            const string inputCode =
                @"Property Let Foo(value)
End Property";

            Assert.AreEqual(1, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void WriteOnlyProperty_ReturnsResult_Set()
        {
            const string inputCode =
                @"Property Set Foo(value)
End Property";

            Assert.AreEqual(1, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void WriteOnlyProperty_ReturnsResult_LetAndSet()
        {
            const string inputCode =
                @"Property Let Foo(value)
End Property

Property Set Foo(value)
End Property";

            Assert.AreEqual(1, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void WriteOnlyProperty_DoesNotReturnsResult_Get()
        {
            const string inputCode =
                @"Property Get Foo()
End Property";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void WriteOnlyProperty_DoesNotReturnsResult_GetAndLetAndSet()
        {
            const string inputCode =
                @"Property Get Foo()
End Property

Property Let Foo(value)
End Property

Property Set Foo(value)
End Property";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void WriteOnlyProperty_ReturnsResult_GetInOtherModule()
        {
            const string classCode =
                @"Property Let Foo(value)
End Property";

            const string otherClassCode =
                @"Property Get Foo()
End Property";

            var inspectionResults = InspectionResultsForModules(
                ("MyClass", classCode, ComponentType.ClassModule),
                ("MyOtherClass", otherClassCode, ComponentType.ClassModule));

            Assert.AreEqual(1, inspectionResults.Count());
        }

        [Test]
        [Category("Inspections")]
        public void WriteOnlyProperty_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore WriteOnlyProperty
Property Let Foo(value)
End Property";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new WriteOnlyPropertyInspection(null);

            Assert.AreEqual(nameof(WriteOnlyPropertyInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new WriteOnlyPropertyInspection(state);
        }
    }
}

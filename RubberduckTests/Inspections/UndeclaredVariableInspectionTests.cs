using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class UndeclaredVariableInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void UndeclaredVariable_ReturnsResult()
        {
            const string inputCode =
                @"Sub Test()
    a = 42
    Debug.Print a
End Sub";

            Assert.AreEqual(1, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule), ReferenceLibrary.VBA).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UndeclaredVariable_ReturnsNoResultIfDeclaredLocally()
        {
            const string inputCode =
                @"Sub Test()
    Dim a As Long
    a = 42
    Debug.Print a
End Sub";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule), ReferenceLibrary.VBA).Count());
        }

        [Test]
        [Category("Inspections")]
        public void UndeclaredVariable_ReturnsNoResultIfDeclaredModuleScope()
        {
            const string inputCode =
                @"Private a As Long
            
Sub Test()
    a = 42
    Debug.Print a
End Sub";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule), ReferenceLibrary.VBA).Count());
        }

        [Test]
        [Category("Inspections")]
        //ReDim acts as a declaration if the array is not declared already.
        //See issue #2522 at https://github.com/rubberduck-vba/Rubberduck/issues/2522
        public void UndeclaredVariable_ReturnsNoResultForReDim()
        {
            const string inputCode =
                @"
Sub Test()
    Dim bar As Variant
    ReDim arr(1 To 42) 
    bar = arr
End Sub";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule)).Count());
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/2525
        [Test]
        [Category("Inspections")]
        public void UndeclaredVariable_ReturnsNoResultIfAnnotated()
        {
            const string inputCode =
                @"Sub Test()
    '@Ignore UndeclaredVariable
    a = 42
    Debug.Print a
End Sub";

            Assert.AreEqual(0, InspectionResultsForModules(("MyClass", inputCode, ComponentType.ClassModule), ReferenceLibrary.VBA).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new UndeclaredVariableInspection(null);

            Assert.AreEqual(nameof(UndeclaredVariableInspection), inspection.Name);
        }


        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new UndeclaredVariableInspection(state);
        }
    }
}

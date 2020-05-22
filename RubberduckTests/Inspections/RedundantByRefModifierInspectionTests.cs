using System.Linq;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Inspections
{
    [TestFixture]
    public class RedundantByRefModifierInspectionTests : InspectionTestsBase
    {
        [Test]
        [Category("Inspections")]
        public void RedundantByRefModifier_ReturnsResult()
        {
            const string inputCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void RedundantByRefModifier_ReturnsResult_MultipleParameters()
        {
            const string inputCode =
                @"Sub Foo(ByRef arg1 As Integer, ByRef arg2 As Date)
End Sub";

            Assert.AreEqual(2, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void RedundantByRefModifier_DoesNotReturnResult_NoModifier()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void RedundantByRefModifier_DoesNotReturnResult_ByVal()
        {
            const string inputCode =
                @"Sub Foo(ByVal arg1 As Integer)
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void RedundantByRefModifier_ReturnsResult_SomePassedByRef()
        {
            const string inputCode =
                @"Sub Foo(ByVal arg1 As Integer, ByRef arg2 As Date)
End Sub";

            Assert.AreEqual(1, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void RedundantByRefModifier_ReturnsResult_InterfaceImplementation()
        {
            const string inputCode1 =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule)
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void RedundantByRefModifier_ReturnsResult_MultipleInterfaceImplementation()
        {
            const string inputCode1 =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode2 =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            const string inputCode3 =
                @"Implements IClass1

Sub IClass1_Foo(ByRef arg1 As Integer)
End Sub";

            var modules = new(string, string, ComponentType)[]
            {
                ("IClass1", inputCode1, ComponentType.ClassModule),
                ("Class1", inputCode2, ComponentType.ClassModule),
                ("Class2", inputCode3, ComponentType.ClassModule)
            };

            Assert.AreEqual(1, InspectionResultsForModules(modules).Count());
        }

        [Test]
        [Category("Inspections")]
        public void RedundantByRefModifier_Ignored_DoesNotReturnResult()
        {
            const string inputCode =
                @"'@Ignore RedundantByRefModifier
Sub Foo(ByRef arg1 As Integer)
End Sub";

            Assert.AreEqual(0, InspectionResultsForStandardModule(inputCode).Count());
        }

        [Test]
        [Category("Inspections")]
        public void InspectionName()
        {
            var inspection = new RedundantByRefModifierInspection(null);

            Assert.AreEqual(nameof(RedundantByRefModifierInspection), inspection.Name);
        }

        protected override IInspection InspectionUnderTest(RubberduckParserState state)
        {
            return new RedundantByRefModifierInspection(state);
        }
    }
}

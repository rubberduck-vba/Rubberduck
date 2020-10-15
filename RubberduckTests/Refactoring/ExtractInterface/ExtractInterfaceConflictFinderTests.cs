using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractInterface;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using System.Linq;

namespace RubberduckTests.Refactoring.ExtractInterface
{
    [TestFixture]
    public class ExtractInterfaceConflictFinderTests
    {
        [TestCase("ITestModule1", true)] //default interfaceName
        [TestCase("TestType", false)] //Public UDT - conflicts
        [TestCase("TestEnum", false)] //Public Enum - conflicts
        [TestCase("TestType2", true)] //Private UDT - OK
        [TestCase("TestEnum2", true)] //Private Enum - OK
        [TestCase("AnotherModule", false)] //Module Identifier - conflicts
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ConflictFinder_IsNonConflictInterfaceName(string interfaceName, bool expectedResult)
        {
            var testModuleCode =
@"Option Explicit

Public Sub MySub()
End Sub";

            var otherModuleName = "AnotherModule";
            var otherModuleCode =
@"Option Explicit

Public Type TestType
    FirstMember As Long
End Type

Public Enum TestEnum
    FirstEnum
End Enum

Private Type TestType2
    FirstMember As Long
End Type

Private Enum TestEnum2
    FirstEnum
End Enum
";

            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, testModuleCode, ComponentType.ClassModule),
                (otherModuleName, otherModuleCode, ComponentType.StandardModule));

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var module = state.DeclarationFinder.MatchName(MockVbeBuilder.TestModuleName).OfType<ClassModuleDeclaration>().Single();

                var conflictFinder = new TestConflictFinderFactory().Create(state, module.ProjectId);                

                Assert.AreEqual(expectedResult, !conflictFinder.IsConflictingModuleName(interfaceName));
            }
        }

        [Test]
        [Category("Refactorings")]
        [Category("Extract Interface")]
        public void ConflictFinder_GenerateNoConflictName()
        {
            var testModuleName = "TargetModule";
            var testModuleCode =
@"Option Explicit

Public Sub MySub()
End Sub";

            var otherModuleName = testModuleName;
            var otherModuleCode =
@"Option Explicit
";

            var vbe = MockVbeBuilder.BuildFromModules((MockVbeBuilder.TestModuleName, testModuleCode, ComponentType.ClassModule),
                (otherModuleName, otherModuleCode, ComponentType.StandardModule));

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var module = state.DeclarationFinder.MatchName(MockVbeBuilder.TestModuleName).OfType<ClassModuleDeclaration>().Single();
                var conflictFinder = new TestConflictFinderFactory().Create(state, module.ProjectId);

                var interfaceModuleName = conflictFinder.GenerateNoConflictModuleName(otherModuleName);

                Assert.AreEqual($"{testModuleName}1", interfaceModuleName);
            }
        }

        private class TestConflictFinderFactory : IExtractInterfaceConflictFinderFactory
        {
            public IExtractInterfaceConflictFinder Create(IDeclarationFinderProvider declarationFinderProvider, string projectId)
            {
                return new ExtractInterfaceConflictFinder(declarationFinderProvider, projectId);
            }
        }
    }
}

using NUnit.Framework;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Linq;
using System.Collections.Generic;
using Rubberduck.Parsing.VBA.ReferenceManagement;

namespace RubberduckTests.Parsing.Coordination
{
    [TestFixture]
    public abstract class IModuleToModuleReferenceManagerTestBase
    {

        protected abstract IModuleToModuleReferenceManager GetNewTestModuleToModuleReferenceManager();

        protected QualifiedModuleName TestModule(string projectName, string moduleName)
        {
            return new QualifiedModuleName(projectName, "dummyPath", moduleName);
        }

        protected QualifiedModuleName TestModule(string moduleName)
        {
            return new QualifiedModuleName("testProject", "dummyPath", moduleName);
        }

        //This property exists to facilitate the setup of concrete implementations of IModuleToModuleReferenceManager.
        protected List<QualifiedModuleName> ModulesUsedInTheBaseTests
        { get
            {
                var usedModules = new List<QualifiedModuleName>();
                usedModules.Add(TestModule("module0"));
                usedModules.Add(TestModule("module1"));
                usedModules.Add(TestModule("module2"));
                usedModules.Add(TestModule("module3"));
                usedModules.Add(TestModule("module4"));
                usedModules.Add(TestModule("module5"));
                return usedModules;
            }
        }


        //Initial Condition Tests

        [Test]
        [Category("Parser")]
        public void ModulesReferencingStartsEmpty()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var testModule = ModulesUsedInTheBaseTests[0];
            var referencingModules = manager.ModulesReferencing(testModule);

            Assert.IsFalse(referencingModules.Any());
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencingAnyStartsEmpty()
        {

            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingModules = manager.ModulesReferencingAny(ModulesUsedInTheBaseTests);

            Assert.IsFalse(referencingModules.Any());
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByStartsEmpty()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var testModule = ModulesUsedInTheBaseTests[0];
            var referencedModules = manager.ModulesReferencedBy(testModule);

            Assert.IsFalse(referencedModules.Any());
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByAnyStartsEmpty()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedModules = manager.ModulesReferencedByAny(ModulesUsedInTheBaseTests);

            Assert.IsFalse(referencedModules.Any());
        }


        //Add Tests

        [Test]
        [Category("Parser")]
        public void ModulesReferencingReturnsAddedReferencesWithMatchingReferencedSide_Single()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.AreEqual(1,referencingModules.Count());
            Assert.IsTrue(referencingModules.Contains(referencingTestModule));
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByReturnsAddedReferencesWithMatchingReferencingModule_Single()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.AreEqual(1,referencedModules.Count());
            Assert.IsTrue(referencedModules.Contains(referencedTestModule));
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencingReturnsAddedReferencesWithMatchingReferencedSide_MultipleDifferent()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.AreEqual(2, referencingModules.Count());
            Assert.IsTrue(referencingModules.Contains(referencingTestModule1));
            Assert.IsTrue(referencingModules.Contains(referencingTestModule2));
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByReturnsAddedReferencesWithMatchingReferencingModule_MultipleDifferent()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule2);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.AreEqual(2, referencedModules.Count());
            Assert.IsTrue(referencedModules.Contains(referencedTestModule1));
            Assert.IsTrue(referencedModules.Contains(referencedTestModule2));
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencingReturnsUniqueValues()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.AreEqual(1,referencingModules.Count());
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByReturnsUniqueValues()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.AreEqual(1, referencedModules.Count());
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencingDoesNotReturnAddedReferencesWithNonMatchingReferencedSide_NoneMatching()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule = ModulesUsedInTheBaseTests[1];
            var notReferencedTestModule = referencingTestModule;
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);

            var referencingModules = manager.ModulesReferencing(notReferencedTestModule);

            Assert.IsFalse(referencingModules.Any());
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByDoesNotReturnAddedReferencesWithNonMatchingReferencingModule_NoneMatching()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule = ModulesUsedInTheBaseTests[1];
            var notReferencingModule = referencedTestModule;
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);

            var referencedModules = manager.ModulesReferencedBy(notReferencingModule);

            Assert.IsFalse(referencedModules.Any());
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencingDoesNotReturnAddedReferencesWithNonMatchingReferencedSide_SomeNotMatching()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[3];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule2);

            var referencingModules = manager.ModulesReferencing(referencedTestModule1);

            Assert.AreEqual(1, referencingModules.Count());
            Assert.IsFalse(referencingModules.Contains(referencingTestModule2));
            Assert.IsTrue(referencingModules.Contains(referencingTestModule1));
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByDoesNotReturnAddedReferencesWithNonMatchingReferencingModule_SomeNotMatching()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[3];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule2);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule1);

            Assert.AreEqual(1, referencedModules.Count());
            Assert.IsFalse(referencedModules.Contains(referencedTestModule2));
            Assert.IsTrue(referencedModules.Contains(referencedTestModule1));
        }


        //Any Tests

        [Test]
        [Category("Parser")]
        public void ModulesReferencingAnyReturnsTheUnionOfTheResultsOfModulesReferencingForTheIndividualModules()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule3 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[3];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[4];
            var referencedTestModule3 = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule2);
            manager.AddModuleToModuleReference(referencingTestModule3, referencedTestModule3);
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule2);

            var referencedModules = new List<QualifiedModuleName> { referencedTestModule1, referencedTestModule2 };
            var referencingModules = manager.ModulesReferencingAny(referencedModules);

            Assert.AreEqual(2, referencingModules.Count());
            Assert.IsTrue(referencingModules.Contains(referencingTestModule1));
            Assert.IsTrue(referencingModules.Contains(referencingTestModule2));
            Assert.IsFalse(referencingModules.Contains(referencingTestModule3));
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByAnyReturnsTheUnionOfTheResultsOfModulesReferencingForTheIndividualModules()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule3 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[3];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[4];
            var referencedTestModule3 = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule2);
            manager.AddModuleToModuleReference(referencingTestModule3, referencedTestModule3);
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule2);

            var referencingModules = new List<QualifiedModuleName> { referencingTestModule1, referencingTestModule2 };
            var referencedModules = manager.ModulesReferencedByAny(referencingModules);

            Assert.AreEqual(2, referencedModules.Count());
            Assert.IsTrue(referencedModules.Contains(referencedTestModule1));
            Assert.IsTrue(referencedModules.Contains(referencedTestModule2));
            Assert.IsFalse(referencedModules.Contains(referencedTestModule3));
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencingAnyReturnsAnEmptyCollectionForEmptyInputCollections()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule3 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[3];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[4];
            var referencedTestModule3 = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule2);
            manager.AddModuleToModuleReference(referencingTestModule3, referencedTestModule3);
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule2);

            var referencedModules = new List<QualifiedModuleName>();
            var referencingModules = manager.ModulesReferencingAny(referencedModules);

            Assert.IsFalse(referencingModules.Any());
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByAnyReturnsAnEmptyCollectionForEmptyInputCollections()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule3 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[3];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[4];
            var referencedTestModule3 = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule2);
            manager.AddModuleToModuleReference(referencingTestModule3, referencedTestModule3);
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule2);

            var referencingModules = new List<QualifiedModuleName>();
            var referencedModules = manager.ModulesReferencedByAny(referencingModules);

            Assert.IsFalse(referencedModules.Any());
        }


        //Remove Tests

        [Test]
        [Category("Parser")]
        public void ModulesReferencingDoesNotReturnResultsForModuleToModuleReferencesThatHaveBeenRemoved()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.IsFalse(referencingModules.Any());
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByDoesNotReturnResultsForModuleToModuleReferencesThatHaveBeenRemoved()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.IsFalse(referencedModules.Any());
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencingReturnsResultsForModuleToModuleReferencesThatHaveNotBeenRemoved_DifferentReferenced()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule, otherReferencedTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencingModules = manager.ModulesReferencing(otherReferencedTestModule);

            Assert.AreEqual(1, referencingModules.Count());
            Assert.IsTrue(referencingModules.Contains(referencingTestModule));
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByReturnsResultsForModuleToModuleReferencesThatHaveNotBeenRemoved_DifferentReferencing()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            manager.AddModuleToModuleReference(otherReferencingTestModule, referencedTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(otherReferencingTestModule);

            Assert.AreEqual(1, referencedModules.Count());
            Assert.IsTrue(referencedModules.Contains(referencedTestModule));
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencingReturnsResultsForModuleToModuleReferencesThatHaveNotBeenRemoved_SameReferenced()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            manager.AddModuleToModuleReference(otherReferencingTestModule, referencedTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.AreEqual(1, referencingModules.Count());
            Assert.IsTrue(referencingModules.Contains(otherReferencingTestModule));
        }

        [Test]
        [Category("Parser")]
        public void ModulesReferencedByReturnsResultsForModuleToModuleReferencesThatHaveNotBeenRemoved_SameReferencing()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule, otherReferencedTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.AreEqual(1, referencedModules.Count());
            Assert.IsTrue(referencedModules.Contains(otherReferencedTestModule));
        }


        //Clear Tests

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesFromModuleRemovesAllMtMReferencesWithTheModuleAsReferencingSide_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule2);
            manager.ClearModuleToModuleReferencesFromModule(referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.IsFalse(referencedModules.Any());
        }

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesFromModuleRemovesAllMtMReferencesWithTheModuleAsReferencingSide_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule2);
            manager.ClearModuleToModuleReferencesFromModule(referencingTestModule);

            var referencingModules1 = manager.ModulesReferencing(referencedTestModule1);
            var referencingModules2 = manager.ModulesReferencing(referencedTestModule2);

            Assert.IsFalse(referencingModules1.Any());
            Assert.IsFalse(referencingModules2.Any());
        }

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesToModuleRemovesAllMtMReferencesWithTheModuleAsReferencedSide_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule);
            manager.ClearModuleToModuleReferencesToModule(referencedTestModule);

            var referencedModules1 = manager.ModulesReferencedBy(referencingTestModule1);
            var referencedModules2 = manager.ModulesReferencedBy(referencingTestModule2);

            Assert.IsFalse(referencedModules1.Any());
            Assert.IsFalse(referencedModules2.Any());
        }

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesToModuleRemovesAllMtMReferencesWithTheModuleAsReferencedSide_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule);
            manager.ClearModuleToModuleReferencesToModule(referencedTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.IsFalse(referencingModules.Any());
        }

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesFromModuleDoesNotRemoveMtMReferencesNotWithTheModuleAsReferencingSide_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[2];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[3];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[4];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule2);
            manager.AddModuleToModuleReference(otherReferencingTestModule, referencedTestModule2);
            manager.AddModuleToModuleReference(otherReferencingTestModule, otherReferencedTestModule);
            manager.ClearModuleToModuleReferencesFromModule(referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(otherReferencingTestModule);

            Assert.AreEqual(2, referencedModules.Count());
            Assert.IsTrue(referencedModules.Contains(referencedTestModule2));
            Assert.IsTrue(referencedModules.Contains(otherReferencedTestModule));
        }

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesFromModuleDoesNotRemoveMtMReferencesNotWithTheModuleAsReferencingSide_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[2];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[3];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[4];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule2);
            manager.AddModuleToModuleReference(otherReferencingTestModule, referencedTestModule2);
            manager.AddModuleToModuleReference(otherReferencingTestModule, otherReferencedTestModule);
            manager.ClearModuleToModuleReferencesFromModule(referencingTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule2);
            var otherReferencingModules = manager.ModulesReferencing(otherReferencedTestModule);

            Assert.AreEqual(1, referencingModules.Count());
            Assert.IsTrue(referencingModules.Contains(otherReferencingTestModule));
            Assert.AreEqual(1, otherReferencingModules.Count());
            Assert.IsTrue(otherReferencingModules.Contains(otherReferencingTestModule));
        }

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesToModuleDoesNotRemoveMtMReferencesNotWithTheModuleAsReferencedSide_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[2];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[3];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[4];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule2, otherReferencedTestModule);
            manager.AddModuleToModuleReference(otherReferencingTestModule, otherReferencedTestModule);
            manager.ClearModuleToModuleReferencesToModule(referencedTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule2);
            var otherReferencedModules = manager.ModulesReferencedBy(otherReferencingTestModule);

            Assert.AreEqual(1, referencedModules.Count());
            Assert.IsTrue(referencedModules.Contains(otherReferencedTestModule));
            Assert.AreEqual(1, otherReferencedModules.Count());
            Assert.IsTrue(otherReferencedModules.Contains(otherReferencedTestModule));
        }

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesToModuleDoesNotRemoveMtMReferencesNotWithTheModuleAsReferencedSide_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[2];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[3];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[4];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule);
            manager.AddModuleToModuleReference(referencingTestModule2, otherReferencedTestModule);
            manager.AddModuleToModuleReference(otherReferencingTestModule, otherReferencedTestModule);
            manager.ClearModuleToModuleReferencesToModule(referencedTestModule);

            var referencingModules = manager.ModulesReferencing(otherReferencedTestModule);

            Assert.AreEqual(2, referencingModules.Count());
            Assert.IsTrue(referencingModules.Contains(referencingTestModule2));
            Assert.IsTrue(referencingModules.Contains(otherReferencingTestModule));
        }


        //Clear Enumerable Overload Tests

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesFromModuleForEnumerablesWorksLikeTheSingleVersionForAllMembersOfTheEnumerable_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[3];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[4];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule2);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule1);
            manager.AddModuleToModuleReference(otherReferencingTestModule, referencedTestModule2);
            manager.AddModuleToModuleReference(otherReferencingTestModule, otherReferencedTestModule);
            var modulesToClearFor = new List<QualifiedModuleName> { referencingTestModule1, referencingTestModule2 };
            manager.ClearModuleToModuleReferencesFromModule(modulesToClearFor);

            var otherReferencedModules = manager.ModulesReferencedBy(otherReferencingTestModule);
            var referencedModules = manager.ModulesReferencedByAny(modulesToClearFor);

            Assert.IsFalse(referencedModules.Any());
            Assert.AreEqual(2, otherReferencedModules.Count());
            Assert.IsTrue(otherReferencedModules.Contains(referencedTestModule2));
            Assert.IsTrue(otherReferencedModules.Contains(otherReferencedTestModule));
        }

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesFromModuleForEnumerablesWorksLikeTheSingleVersionForAllMembersOfTheEnumerable_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[3];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[4];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule2);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule1);
            manager.AddModuleToModuleReference(otherReferencingTestModule, referencedTestModule2);
            manager.AddModuleToModuleReference(otherReferencingTestModule, otherReferencedTestModule);
            var modulesToClearFor = new List<QualifiedModuleName> { referencingTestModule1, referencingTestModule2 };
            manager.ClearModuleToModuleReferencesFromModule(modulesToClearFor);

            var referencingModules1 = manager.ModulesReferencing(referencedTestModule1);
            var referencingModules2 = manager.ModulesReferencing(referencedTestModule2);
            var otherReferencingModules = manager.ModulesReferencing(otherReferencedTestModule);

            Assert.IsFalse(referencingModules1.Any());
            Assert.AreEqual(1, referencingModules2.Count());
            Assert.IsTrue(referencingModules2.Contains(otherReferencingTestModule));
            Assert.AreEqual(1, otherReferencingModules.Count());
            Assert.IsTrue(otherReferencingModules.Contains(otherReferencingTestModule));
        }


        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesToModuleForEnumerablesWorksLikeTheSingleVersionForAllMembersOfTheEnumerable_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[3];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[4];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule2);
            manager.AddModuleToModuleReference(referencingTestModule2, otherReferencedTestModule);
            manager.AddModuleToModuleReference(otherReferencingTestModule, otherReferencedTestModule);
            var modulesToClearFor = new List<QualifiedModuleName> { referencedTestModule1, referencedTestModule2 };
            manager.ClearModuleToModuleReferencesToModule(modulesToClearFor);

            var referencedModules1 = manager.ModulesReferencedBy(referencingTestModule1);
            var referencedModules2 = manager.ModulesReferencedBy(referencingTestModule2);
            var otherReferencedModules = manager.ModulesReferencedBy(otherReferencingTestModule);

            Assert.IsFalse(referencedModules1.Any());
            Assert.AreEqual(1, referencedModules2.Count());
            Assert.IsTrue(referencedModules2.Contains(otherReferencedTestModule));
            Assert.AreEqual(1, otherReferencedModules.Count());
            Assert.IsTrue(otherReferencedModules.Contains(otherReferencedTestModule));
        }

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesToModuleForEnumerablesWorksLikeTheSingleVersionForAllMembersOfTheEnumerable_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[3];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[4];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule2, referencedTestModule1);
            manager.AddModuleToModuleReference(referencingTestModule1, referencedTestModule2);
            manager.AddModuleToModuleReference(referencingTestModule2, otherReferencedTestModule);
            manager.AddModuleToModuleReference(otherReferencingTestModule, otherReferencedTestModule);
            var modulesToClearFor = new List<QualifiedModuleName> { referencedTestModule1, referencedTestModule2 };
            manager.ClearModuleToModuleReferencesToModule(modulesToClearFor);

            var otherReferencingModules = manager.ModulesReferencing(otherReferencedTestModule);
            var referencingModules = manager.ModulesReferencingAny(modulesToClearFor);

            Assert.IsFalse(referencingModules.Any());
            Assert.AreEqual(2, otherReferencingModules.Count());
            Assert.IsTrue(otherReferencingModules.Contains(referencingTestModule2));
            Assert.IsTrue(otherReferencingModules.Contains(otherReferencingTestModule));
        }

        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesFromModuleDoesNothingForAnEmptyEnumerables()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            var modulesToClearFor = new List<QualifiedModuleName>();
            manager.ClearModuleToModuleReferencesFromModule(modulesToClearFor);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);
            var referencedgModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.AreEqual(1, referencingModules.Count());
            Assert.IsTrue(referencingModules.Contains(referencingTestModule));
            Assert.AreEqual(1, referencedgModules.Count());
            Assert.IsTrue(referencedgModules.Contains(referencedTestModule));
        }


        [Test]
        [Category("Parser")]
        public void ClearMtMReferencesToModuleDoesNothingForAnEmptyEnumerables()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencingTestModule, referencedTestModule);
            var modulesToClearFor = new List<QualifiedModuleName>();
            manager.ClearModuleToModuleReferencesToModule(modulesToClearFor);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);
            var referencedgModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.AreEqual(1, referencingModules.Count());
            Assert.IsTrue(referencingModules.Contains(referencingTestModule));
            Assert.AreEqual(1, referencedgModules.Count());
            Assert.IsTrue(referencedgModules.Contains(referencedTestModule));
        }
    }
}

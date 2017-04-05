using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Linq;
using System.Collections.Generic;

namespace RubberduckTests.Parsing.Coordination
{
    [TestClass]
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

        [TestMethod]
        public void ModulesReferencingStartsEmpty()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var testModule = ModulesUsedInTheBaseTests[0];
            var referencingModules = manager.ModulesReferencing(testModule);

            Assert.IsFalse(referencingModules.Any());
        }

        [TestMethod]
        public void ModulesReferencingAnyStartsEmpty()
        {

            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingModules = manager.ModulesReferencingAny(ModulesUsedInTheBaseTests);

            Assert.IsFalse(referencingModules.Any());
        }

        [TestMethod]
        public void ModulesReferencedByStartsEmpty()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var testModule = ModulesUsedInTheBaseTests[0];
            var referencedModules = manager.ModulesReferencedBy(testModule);

            Assert.IsFalse(referencedModules.Any());
        }

        [TestMethod]
        public void ModulesReferencedByAnyStartsEmpty()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedModules = manager.ModulesReferencedByAny(ModulesUsedInTheBaseTests);

            Assert.IsFalse(referencedModules.Any());
        }


        //Add Tests

        [TestMethod]
        public void ModulesReferencingReturnsAddedReferencesWithMatchingReferencedSide_Single()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.IsTrue(referencingModules.Count() == 1);
            Assert.IsTrue(referencingModules.Contains(referencingTestModule));
        }

        [TestMethod]
        public void ModulesReferencedByReturnsAddedReferencesWithMatchingReferencingModule_Single()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.IsTrue(referencedModules.Count() == 1);
            Assert.IsTrue(referencedModules.Contains(referencedTestModule));
        }

        [TestMethod]
        public void ModulesReferencingReturnsAddedReferencesWithMatchingReferencedSide_MultipleDifferent()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule2);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.IsTrue(referencingModules.Count() == 2);
            Assert.IsTrue(referencingModules.Contains(referencingTestModule1));
            Assert.IsTrue(referencingModules.Contains(referencingTestModule2));
        }

        [TestMethod]
        public void ModulesReferencedByReturnsAddedReferencesWithMatchingReferencingModule_MultipleDifferent()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.IsTrue(referencedModules.Count() == 2);
            Assert.IsTrue(referencedModules.Contains(referencedTestModule1));
            Assert.IsTrue(referencedModules.Contains(referencedTestModule2));
        }

        [TestMethod]
        public void ModulesReferencingReturnsUniqueValues()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.IsTrue(referencingModules.Count() == 1);
        }

        [TestMethod]
        public void ModulesReferencedByReturnsUniqueValues()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.IsTrue(referencedModules.Count() == 1);
        }

        [TestMethod]
        public void ModulesReferencingDoesNotReturnAddedReferencesWithNonMatchingReferencedSide_NoneMatching()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule = ModulesUsedInTheBaseTests[1];
            var notReferencedTestModule = referencingTestModule;
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencingModules = manager.ModulesReferencing(notReferencedTestModule);

            Assert.IsFalse(referencingModules.Any());
        }

        [TestMethod]
        public void ModulesReferencedByDoesNotReturnAddedReferencesWithNonMatchingReferencingModule_NoneMatching()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule = ModulesUsedInTheBaseTests[1];
            var notReferencingModule = referencedTestModule;
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(notReferencingModule);

            Assert.IsFalse(referencedModules.Any());
        }

        [TestMethod]
        public void ModulesReferencingDoesNotReturnAddedReferencesWithNonMatchingReferencedSide_SomeNotMatching()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[3];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule2);

            var referencingModules = manager.ModulesReferencing(referencedTestModule1);

            Assert.IsTrue(referencingModules.Count() == 1);
            Assert.IsFalse(referencingModules.Contains(referencingTestModule2));
            Assert.IsTrue(referencingModules.Contains(referencingTestModule1));
        }

        [TestMethod]
        public void ModulesReferencedByDoesNotReturnAddedReferencesWithNonMatchingReferencingModule_SomeNotMatching()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[3];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule2);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule1);

            Assert.IsTrue(referencedModules.Count() == 1);
            Assert.IsFalse(referencedModules.Contains(referencedTestModule2));
            Assert.IsTrue(referencedModules.Contains(referencedTestModule1));
        }


        //Any Tests

        [TestMethod]
        public void ModulesReferencingAnyReturnsTheUnionOfTheResultsOfModulesReferencingForTheIndividualModules()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule3 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[3];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[4];
            var referencedTestModule3 = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule2);
            manager.AddModuleToModuleReference(referencedTestModule3, referencingTestModule3);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule1);

            var referencedModules = new List<QualifiedModuleName> { referencedTestModule1, referencedTestModule2 };
            var referencingModules = manager.ModulesReferencingAny(referencedModules);

            Assert.IsTrue(referencingModules.Count() == 2);
            Assert.IsTrue(referencingModules.Contains(referencingTestModule1));
            Assert.IsTrue(referencingModules.Contains(referencingTestModule2));
            Assert.IsFalse(referencingModules.Contains(referencingTestModule3));
        }

        [TestMethod]
        public void ModulesReferencedByAnyReturnsTheUnionOfTheResultsOfModulesReferencingForTheIndividualModules()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule3 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[3];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[4];
            var referencedTestModule3 = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule2);
            manager.AddModuleToModuleReference(referencedTestModule3, referencingTestModule3);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule1);

            var referencingModules = new List<QualifiedModuleName> { referencingTestModule1, referencingTestModule2 };
            var referencedModules = manager.ModulesReferencedByAny(referencingModules);

            Assert.IsTrue(referencedModules.Count() == 2);
            Assert.IsTrue(referencedModules.Contains(referencedTestModule1));
            Assert.IsTrue(referencedModules.Contains(referencedTestModule2));
            Assert.IsFalse(referencedModules.Contains(referencedTestModule3));
        }

        [TestMethod]
        public void ModulesReferencingAnyReturnsAnEmptyCollectionForEmptyInputCollections()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule3 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[3];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[4];
            var referencedTestModule3 = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule2);
            manager.AddModuleToModuleReference(referencedTestModule3, referencingTestModule3);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule1);

            var referencedModules = new List<QualifiedModuleName>();
            var referencingModules = manager.ModulesReferencingAny(referencedModules);

            Assert.IsFalse(referencingModules.Any());
        }

        [TestMethod]
        public void ModulesReferencedByAnyReturnsAnEmptyCollectionForEmptyInputCollections()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule3 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[3];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[4];
            var referencedTestModule3 = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule2);
            manager.AddModuleToModuleReference(referencedTestModule3, referencingTestModule3);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule1);

            var referencingModules = new List<QualifiedModuleName>();
            var referencedModules = manager.ModulesReferencedByAny(referencingModules);

            Assert.IsFalse(referencedModules.Any());
        }


        //Remove Tests

        [TestMethod]
        public void ModulesReferencingDoesNotReturnResultsForModuleToModuleReferencesThatHaveBeenRemoved()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.IsFalse(referencingModules.Any());
        }

        [TestMethod]
        public void ModulesReferencedByDoesNotReturnResultsForModuleToModuleReferencesThatHaveBeenRemoved()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.IsFalse(referencedModules.Any());
        }

        [TestMethod]
        public void ModulesReferencingReturnsResultsForModuleToModuleReferencesThatHaveNotBeenRemoved_DifferentReferenced()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            manager.AddModuleToModuleReference(otherReferencedTestModule, referencingTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencingModules = manager.ModulesReferencing(otherReferencedTestModule);

            Assert.IsTrue(referencingModules.Count() == 1);
            Assert.IsTrue(referencingModules.Contains(referencingTestModule));
        }

        [TestMethod]
        public void ModulesReferencedByReturnsResultsForModuleToModuleReferencesThatHaveNotBeenRemoved_DifferentReferencing()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule, otherReferencingTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(otherReferencingTestModule);

            Assert.IsTrue(referencedModules.Count() == 1);
            Assert.IsTrue(referencedModules.Contains(referencedTestModule));
        }

        [TestMethod]
        public void ModulesReferencingReturnsResultsForModuleToModuleReferencesThatHaveNotBeenRemoved_SameReferenced()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule, otherReferencingTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.IsTrue(referencingModules.Count() == 1);
            Assert.IsTrue(referencingModules.Contains(otherReferencingTestModule));
        }

        [TestMethod]
        public void ModulesReferencedByReturnsResultsForModuleToModuleReferencesThatHaveNotBeenRemoved_SameReferencing()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            manager.AddModuleToModuleReference(otherReferencedTestModule, referencingTestModule);
            manager.RemoveModuleToModuleReference(referencedTestModule, referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.IsTrue(referencedModules.Count() == 1);
            Assert.IsTrue(referencedModules.Contains(otherReferencedTestModule));
        }


        //Clear Tests

        [TestMethod]
        public void ClearMtMReferencesFromModuleRemovesAllMtMReferencesWithTheModuleAsReferencingSide_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule);
            manager.ClearModuleToModuleReferencesFromModule(referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.IsFalse(referencedModules.Any());
        }

        [TestMethod]
        public void ClearMtMReferencesFromModuleRemovesAllMtMReferencesWithTheModuleAsReferencingSide_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule);
            manager.ClearModuleToModuleReferencesFromModule(referencingTestModule);

            var referencingModules1 = manager.ModulesReferencing(referencedTestModule1);
            var referencingModules2 = manager.ModulesReferencing(referencedTestModule2);

            Assert.IsFalse(referencingModules1.Any());
            Assert.IsFalse(referencingModules2.Any());
        }

        [TestMethod]
        public void ClearMtMReferencesToModuleRemovesAllMtMReferencesWithTheModuleAsReferencedSide_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule2);
            manager.ClearModuleToModuleReferencesToModule(referencedTestModule);

            var referencedModules1 = manager.ModulesReferencedBy(referencingTestModule1);
            var referencedModules2 = manager.ModulesReferencedBy(referencingTestModule2);

            Assert.IsFalse(referencedModules1.Any());
            Assert.IsFalse(referencedModules2.Any());
        }

        [TestMethod]
        public void ClearMtMReferencesToModuleRemovesAllMtMReferencesWithTheModuleAsReferencedSide_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[2];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule2);
            manager.ClearModuleToModuleReferencesToModule(referencedTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);

            Assert.IsFalse(referencingModules.Any());
        }

        [TestMethod]
        public void ClearMtMReferencesFromModuleDoesNotRemoveMtMReferencesNotWithTheModuleAsReferencingSide_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[2];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[3];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[4];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule2, otherReferencingTestModule);
            manager.AddModuleToModuleReference(otherReferencedTestModule, otherReferencingTestModule);
            manager.ClearModuleToModuleReferencesFromModule(referencingTestModule);

            var referencedModules = manager.ModulesReferencedBy(otherReferencingTestModule);

            Assert.IsTrue(referencedModules.Count() == 2);
            Assert.IsTrue(referencedModules.Contains(referencedTestModule2));
            Assert.IsTrue(referencedModules.Contains(otherReferencedTestModule));
        }

        [TestMethod]
        public void ClearMtMReferencesFromModuleDoesNotRemoveMtMReferencesNotWithTheModuleAsReferencingSide_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[2];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[3];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[4];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule);
            manager.AddModuleToModuleReference(referencedTestModule2, otherReferencingTestModule);
            manager.AddModuleToModuleReference(otherReferencedTestModule, otherReferencingTestModule);
            manager.ClearModuleToModuleReferencesFromModule(referencingTestModule);

            var referencingModules = manager.ModulesReferencing(referencedTestModule2);
            var otherReferencingModules = manager.ModulesReferencing(otherReferencedTestModule);

            Assert.IsTrue(referencingModules.Count() == 1);
            Assert.IsTrue(referencingModules.Contains(otherReferencingTestModule));
            Assert.IsTrue(otherReferencingModules.Count() == 1);
            Assert.IsTrue(otherReferencingModules.Contains(otherReferencingTestModule));
        }

        [TestMethod]
        public void ClearMtMReferencesToModuleDoesNotRemoveMtMReferencesNotWithTheModuleAsReferencedSide_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[2];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[3];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[4];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule2);
            manager.AddModuleToModuleReference(otherReferencedTestModule, referencingTestModule2);
            manager.AddModuleToModuleReference(otherReferencedTestModule, otherReferencingTestModule);
            manager.ClearModuleToModuleReferencesToModule(referencedTestModule);

            var referencedModules = manager.ModulesReferencedBy(referencingTestModule2);
            var otherReferencedModules = manager.ModulesReferencedBy(otherReferencingTestModule);

            Assert.IsTrue(referencedModules.Count() == 1);
            Assert.IsTrue(referencedModules.Contains(otherReferencedTestModule));
            Assert.IsTrue(otherReferencedModules.Count() == 1);
            Assert.IsTrue(otherReferencedModules.Contains(otherReferencedTestModule));
        }

        [TestMethod]
        public void ClearMtMReferencesToModuleDoesNotRemoveMtMReferencesNotWithTheModuleAsReferencedSide_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule = ModulesUsedInTheBaseTests[0];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[2];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[3];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[4];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule2);
            manager.AddModuleToModuleReference(otherReferencedTestModule, referencingTestModule2);
            manager.AddModuleToModuleReference(otherReferencedTestModule, otherReferencingTestModule);
            manager.ClearModuleToModuleReferencesToModule(referencedTestModule);

            var referencingModules = manager.ModulesReferencing(otherReferencedTestModule);

            Assert.IsTrue(referencingModules.Count() == 2);
            Assert.IsTrue(referencingModules.Contains(referencingTestModule2));
            Assert.IsTrue(referencingModules.Contains(otherReferencingTestModule));
        }


        //Clear Enumerable Overload Tests

        [TestMethod]
        public void ClearMtMReferencesFromModuleForEnumerablesWorksLikeTheSingleVersionForAllMembersOfTheEnumerable_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[3];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[4];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule2);
            manager.AddModuleToModuleReference(referencedTestModule2, otherReferencingTestModule);
            manager.AddModuleToModuleReference(otherReferencedTestModule, otherReferencingTestModule);
            var modulesToClearFor = new List<QualifiedModuleName> { referencingTestModule1, referencingTestModule2 };
            manager.ClearModuleToModuleReferencesFromModule(modulesToClearFor);

            var otherReferencedModules = manager.ModulesReferencedBy(otherReferencingTestModule);
            var referencedModules = manager.ModulesReferencedByAny(modulesToClearFor);

            Assert.IsFalse(referencedModules.Any());
            Assert.IsTrue(otherReferencedModules.Count() == 2);
            Assert.IsTrue(otherReferencedModules.Contains(referencedTestModule2));
            Assert.IsTrue(otherReferencedModules.Contains(otherReferencedTestModule));
        }

        [TestMethod]
        public void ClearMtMReferencesFromModuleForEnumerablesWorksLikeTheSingleVersionForAllMembersOfTheEnumerable_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencedTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[3];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[4];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule2);
            manager.AddModuleToModuleReference(referencedTestModule2, otherReferencingTestModule);
            manager.AddModuleToModuleReference(otherReferencedTestModule, otherReferencingTestModule);
            var modulesToClearFor = new List<QualifiedModuleName> { referencingTestModule1, referencingTestModule2 };
            manager.ClearModuleToModuleReferencesFromModule(modulesToClearFor);

            var referencingModules1 = manager.ModulesReferencing(referencedTestModule1);
            var referencingModules2 = manager.ModulesReferencing(referencedTestModule2);
            var otherReferencingModules = manager.ModulesReferencing(otherReferencedTestModule);

            Assert.IsFalse(referencingModules1.Any());
            Assert.IsTrue(referencingModules2.Count() == 1);
            Assert.IsTrue(referencingModules2.Contains(otherReferencingTestModule));
            Assert.IsTrue(otherReferencingModules.Count() == 1);
            Assert.IsTrue(otherReferencingModules.Contains(otherReferencingTestModule));
        }


        [TestMethod]
        public void ClearMtMReferencesToModuleForEnumerablesWorksLikeTheSingleVersionForAllMembersOfTheEnumerable_ReferencingSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[3];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[4];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule2);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule1);
            manager.AddModuleToModuleReference(otherReferencedTestModule, referencingTestModule2);
            manager.AddModuleToModuleReference(otherReferencedTestModule, otherReferencingTestModule);
            var modulesToClearFor = new List<QualifiedModuleName> { referencedTestModule1, referencedTestModule2 };
            manager.ClearModuleToModuleReferencesToModule(modulesToClearFor);

            var referencedModules1 = manager.ModulesReferencedBy(referencingTestModule2);
            var referencedModules2 = manager.ModulesReferencedBy(referencingTestModule2);
            var otherReferencedModules = manager.ModulesReferencedBy(otherReferencingTestModule);

            Assert.IsFalse(referencedModules1.Any());
            Assert.IsTrue(referencedModules2.Count() == 1);
            Assert.IsTrue(referencedModules2.Contains(otherReferencedTestModule));
            Assert.IsTrue(otherReferencedModules.Count() == 1);
            Assert.IsTrue(otherReferencedModules.Contains(otherReferencedTestModule));
        }

        [TestMethod]
        public void ClearMtMReferencesToModuleForEnumerablesWorksLikeTheSingleVersionForAllMembersOfTheEnumerable_ReferncedSide()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencedTestModule1 = ModulesUsedInTheBaseTests[0];
            var referencedTestModule2 = ModulesUsedInTheBaseTests[1];
            var referencingTestModule1 = ModulesUsedInTheBaseTests[2];
            var referencingTestModule2 = ModulesUsedInTheBaseTests[3];
            var otherReferencingTestModule = ModulesUsedInTheBaseTests[4];
            var otherReferencedTestModule = ModulesUsedInTheBaseTests[5];
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule1);
            manager.AddModuleToModuleReference(referencedTestModule1, referencingTestModule2);
            manager.AddModuleToModuleReference(referencedTestModule2, referencingTestModule1);
            manager.AddModuleToModuleReference(otherReferencedTestModule, referencingTestModule2);
            manager.AddModuleToModuleReference(otherReferencedTestModule, otherReferencingTestModule);
            var modulesToClearFor = new List<QualifiedModuleName> { referencedTestModule1, referencedTestModule2 };
            manager.ClearModuleToModuleReferencesToModule(modulesToClearFor);

            var otherReferencingModules = manager.ModulesReferencing(otherReferencedTestModule);
            var referencingModules = manager.ModulesReferencingAny(modulesToClearFor);

            Assert.IsFalse(referencingModules.Any());
            Assert.IsTrue(otherReferencingModules.Count() == 2);
            Assert.IsTrue(otherReferencingModules.Contains(referencingTestModule2));
            Assert.IsTrue(otherReferencingModules.Contains(otherReferencingTestModule));
        }

        [TestMethod]
        public void ClearMtMReferencesFromModuleDoesNothingForAnEmptyEnumerables()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            var modulesToClearFor = new List<QualifiedModuleName>();
            manager.ClearModuleToModuleReferencesFromModule(modulesToClearFor);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);
            var referencedgModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.IsTrue(referencingModules.Count() == 1);
            Assert.IsTrue(referencingModules.Contains(referencingTestModule));
            Assert.IsTrue(referencedgModules.Count() == 1);
            Assert.IsTrue(referencedgModules.Contains(referencedTestModule));
        }


        [TestMethod]
        public void ClearMtMReferencesToModuleDoesNothingForAnEmptyEnumerables()
        {
            var manager = GetNewTestModuleToModuleReferenceManager();
            var referencingTestModule = ModulesUsedInTheBaseTests[0];
            var referencedTestModule = ModulesUsedInTheBaseTests[1];
            manager.AddModuleToModuleReference(referencedTestModule, referencingTestModule);
            var modulesToClearFor = new List<QualifiedModuleName>();
            manager.ClearModuleToModuleReferencesToModule(modulesToClearFor);

            var referencingModules = manager.ModulesReferencing(referencedTestModule);
            var referencedgModules = manager.ModulesReferencedBy(referencingTestModule);

            Assert.IsTrue(referencingModules.Count() == 1);
            Assert.IsTrue(referencingModules.Contains(referencingTestModule));
            Assert.IsTrue(referencedgModules.Count() == 1);
            Assert.IsTrue(referencedgModules.Contains(referencedTestModule));
        }
    }
}

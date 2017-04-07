using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Symbols
{
    [TestClass]
    public class ClassModuleDeclarationTests
    {
        [TestCategory("Resolver")]
        [TestMethod]
        public void ClassModulesHaveDeclarationTypeClassModule()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsTrue(classModule.DeclarationType.HasFlag(DeclarationType.ClassModule));
        }

            private static ProjectDeclaration GetTestProject(string name)
            {
                var qualifiedProjectName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ProjectDeclaration(qualifiedProjectName, name, true, null);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }

            private static ClassModuleDeclaration GetTestClassModule(Declaration projectDeclatation, string name, bool isUserDefined, Attributes attributes, bool hasDefaultInstanceVariable = false)
            {
                var qualifiedClassModuleMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ClassModuleDeclaration(qualifiedClassModuleMemberName, projectDeclatation, name, isUserDefined, null, attributes, hasDefaultInstanceVariable);
            }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ByDefaultSubtypesIsEmpty()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.Subtypes.Any());
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void AddSubtypeAddsClassToSubtypes()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var subtype = GetTestClassModule(projectDeclaration, "testSubtype", true, null);
            classModule.AddSubtype(subtype);

            Assert.IsTrue(classModule.Subtypes.First().Equals(subtype));
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ByDefaultSupertypesIsEmpty()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.Supertypes.Any());
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void AddSupertypeForDeclarationsAddsClassToSupertypes()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var supertype = GetTestClassModule(projectDeclaration, "testSupertype", true, null);
            classModule.AddSupertype(supertype);

            Assert.IsTrue(classModule.Supertypes.First().Equals(supertype));
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ByDefaultSupertypeNamesIsEmpty()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.SupertypeNames.Any());
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void AddSupertypeForStringsAddsTypenameToSupertypeNames()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var supertypeName = "testSupertypeName";
            classModule.AddSupertype(supertypeName);

            Assert.IsTrue(classModule.SupertypeNames.First().Equals(supertypeName));
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void AddSupertypeForDeclarationsHasNoEffectOnSupertypeNames()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var supertype = GetTestClassModule(projectDeclaration, "testSupertype", true, null);
            classModule.AddSupertype(supertype);

            Assert.IsFalse(classModule.SupertypeNames.Any());
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void AddSupertypeForStringsHasNoEffectsOnSupertypes()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var supertypeName = "testSupertypeName";
            classModule.AddSupertype(supertypeName);

            Assert.IsFalse(classModule.Supertypes.Any());
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void GetSupertypesReturnsAnEmptyEnumerableForProceduralModules()
        {
            var projectDeclaration = GetTestProject("testProject");
            var proceduralModule = GetTestProceduralModule(projectDeclaration, "testModule");

            Assert.IsFalse(ClassModuleDeclaration.GetSupertypes(proceduralModule).Any());
        }

            private static ProceduralModuleDeclaration GetTestProceduralModule(Declaration projectDeclatation, string name)
            {
                var qualifiedClassModuleMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ProceduralModuleDeclaration(qualifiedClassModuleMemberName, projectDeclatation, name, true, null, null);
            }


        [TestCategory("Resolver")]
        [TestMethod]
        public void GetSupertypesReturnsTheSupertypesOfAClassModule()
        {
            var projectDeclaration = GetTestProject("testProject");
            var supertype = GetTestClassModule(projectDeclaration, "testSupertype", true, null);
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            classModule.AddSupertype(supertype);

            Assert.AreEqual(supertype, ClassModuleDeclaration.GetSupertypes(classModule).First());
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void GetSupertypesReturnsAnEmptyEnumerableForDeclarationsWithDeclarationTypeClassModuleWhichAreNoClassModuleDeclarations()
        {
            var projectDeclaration = GetTestProject("testProject");
            var fakeClassModule = GetTestFakeClassModule(projectDeclaration, "testFakeClass");

            Assert.IsFalse(ClassModuleDeclaration.GetSupertypes(fakeClassModule).Any());
        }

            private static Declaration GetTestFakeClassModule(Declaration parentDeclatation, string name)
            {
                var qualifiedVariableMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new Declaration(qualifiedVariableMemberName, parentDeclatation, "dummy", "test", "test", false, false, Accessibility.Public, DeclarationType.ClassModule, null, Selection.Home, true, null);
            }



        [TestCategory("Resolver")]
        [TestMethod]
        public void ByDefaultDefaultMemberIsNull()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsNull(classModule.DefaultMember);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ByDefaultClassModulesNotBuiltInAreNotExposed()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.IsExposed);
        }


        // TODO: Find out if there's info about "being exposed" in type libraries.
        // We take the conservative approach of treating all type library modules as exposed.
        [TestCategory("Resolver")]
        [TestMethod]
        public void BuiltInClassesAreExposed()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", false, null);

            Assert.IsTrue(classModule.IsExposed);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ClassModulesWithTheExposedAttributeAreExposed()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddExposedClassAttribute();
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, classAttributes);

            Assert.IsTrue(classModule.IsExposed);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ByDefaultClassModulesAreNotGlobalClasses()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.IsGlobalClassModule);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ClassModulesWithTheGlobalNamespaceAttributeAreGlobalClasses()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddGlobalClassAttribute();
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, classAttributes);

            Assert.IsTrue(classModule.IsGlobalClassModule);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ClassModulesWithASubtypeBelowInTheHiearchyThatIsAGlobalClassAndThatHasBeenAddedBeforeCallingIsGlobalClassTheFirstTimeIsAGlobalClass()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddGlobalClassAttribute();
            var subsubtype = GetTestClassModule(projectDeclaration, "testSubSubtype", true, classAttributes);
            var subtype = GetTestClassModule(projectDeclaration, "testSubtype", true, null);
            subtype.AddSubtype(subsubtype);
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            classModule.AddSubtype(subtype);

            Assert.IsTrue(classModule.IsGlobalClassModule);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ClassModulesDoNotBecomeAGlobalClassIfASubtypeBelowInTheHiearchyIsAddedThatIsAGlobalClassAfterIsAGlobalClassHasAlreadyBeenCalled()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddGlobalClassAttribute();
            var subsubtype = GetTestClassModule(projectDeclaration, "testSubSubtype", true, classAttributes);
            var subtype = GetTestClassModule(projectDeclaration, "testSubtype", true, null);
            subtype.AddSubtype(subsubtype);
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var dummy = classModule.IsGlobalClassModule;
            classModule.AddSubtype(subtype);

            Assert.IsFalse(classModule.IsGlobalClassModule);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ClassModulesDoNotBecomeAGlobalClassIfBelowInTheHierarchyASubtypeIsAddedThatIsAGlobalClassAfterIsAGlobalClassHasAlreadyBeenCalled()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddGlobalClassAttribute();
            var subsubtype = GetTestClassModule(projectDeclaration, "testSubSubtype", true, classAttributes);
            var subtype = GetTestClassModule(projectDeclaration, "testSubtype", true, null);
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            classModule.AddSubtype(subtype);
            var dummy = classModule.IsGlobalClassModule;
            subtype.AddSubtype(subsubtype);

            Assert.IsFalse(classModule.IsGlobalClassModule);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ByDefaultClassModulesDoNotHaveAPredeclaredID()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.HasPredeclaredId);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ClassModulesHaveAPredeclaredIDIfStatedInTheConstructorThatTheyHaveADefaultInstanceVariable()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null, true);

            Assert.IsTrue(classModule.HasPredeclaredId);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ClassModulesWithThePredeclaredIDAttributeHaveAPredeclaredID()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddPredeclaredIdTypeAttribute();
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, classAttributes);

            Assert.IsTrue(classModule.HasPredeclaredId);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ByDefaultClassModulesDoNotHaveADefaultInstanceVariable()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.HasDefaultInstanceVariable);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ClassModulesThatAreGlobalClassesHaveADefaultInstanceVariable()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddGlobalClassAttribute();
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, classAttributes);

            Assert.IsTrue(classModule.HasDefaultInstanceVariable);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ClassModulesWithThePredeclaredIDAttributeHaveADefaultInstanceVariable()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddPredeclaredIdTypeAttribute();
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, classAttributes);

            Assert.IsTrue(classModule.HasDefaultInstanceVariable);
        }


        [TestCategory("Resolver")]
        [TestMethod]
        public void ClassModulesHaveADefaultInstanceVariableIfThisIsStated()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null, true);

            Assert.IsTrue(classModule.HasDefaultInstanceVariable);
        }

    }
}

using NUnit.Framework;
using System.Linq;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class ClassModuleDeclarationTests
    {
        [Category("Resolver")]
        [Test]
        public void ClassModulesHaveDeclarationTypeClassModule()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsTrue(classModule.DeclarationType.HasFlag(DeclarationType.ClassModule));
        }

            private static ProjectDeclaration GetTestProject(string name)
            {
                var qualifiedProjectName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ProjectDeclaration(qualifiedProjectName, name, true);
            }

                private static QualifiedModuleName StubQualifiedModuleName()
                {
                    return new QualifiedModuleName("dummy", "dummy", "dummy");
                }

            private static ClassModuleDeclaration GetTestClassModule(Declaration projectDeclatation, string name, bool isUserDefined, Attributes attributes, bool hasDefaultInstanceVariable = false)
            {
                var qualifiedClassModuleMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new ClassModuleDeclaration(qualifiedClassModuleMemberName, projectDeclatation, name, isUserDefined, null, attributes, false, hasDefaultInstanceVariable);
            }


        [Category("Resolver")]
        [Test]
        public void ByDefaultSubtypesIsEmpty()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.Subtypes.Any());
        }


        [Category("Resolver")]
        [Test]
        public void AddSupertypeAddsClassToSubtypesOfSupertype()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var subtype = GetTestClassModule(projectDeclaration, "testSubtype", true, null);
            subtype.AddSupertype(classModule);

            Assert.IsTrue(classModule.Subtypes.First().Equals(subtype));
        }


        [Category("Resolver")]
        [Test]
        public void ByDefaultSupertypesIsEmpty()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.Supertypes.Any());
        }


        [Category("Resolver")]
        [Test]
        public void AddSupertypeAddsClassToSupertypes()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var supertype = GetTestClassModule(projectDeclaration, "testSupertype", true, null);
            classModule.AddSupertype(supertype);

            Assert.IsTrue(classModule.Supertypes.First().Equals(supertype));
        }

        [Category("Resolver")]
        [Test]
        public void ClearSupertypeRemovesAllSupertypes()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var supertype1 = GetTestClassModule(projectDeclaration, "testSupertype1", true, null);
            var supertype2 = GetTestClassModule(projectDeclaration, "testSupertype2", true, null);
            classModule.AddSupertype(supertype1);
            classModule.AddSupertype(supertype2);
            classModule.ClearSupertypes();

            Assert.IsFalse(classModule.Supertypes.Any());
        }

        //The reasoning behind this is that the names of the supertypes only depend on the module itself.
        //So, the module itself has to be changed to change them. That in turn would mean a reparse and discarding the module declaration. 
        [Category("Resolver")]
        [Test]
        public void ClearSupertypeDoesNotRemoveSupertypesNames()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            classModule.AddSupertypeName("testSupertype1");
            classModule.AddSupertypeName("testSupertype2");
            classModule.ClearSupertypes();
            var supertypeNameCount = classModule.SupertypeNames.Count();

            Assert.AreEqual(2, supertypeNameCount);
        }

        [Category("Resolver")]
        [Test]
        public void ClearSupertypeRemovesAllSupertypesRemovesTheClassFromTheSubtypesOfTheSupertypes()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var supertype1 = GetTestClassModule(projectDeclaration, "testSupertype1", true, null);
            var supertype2 = GetTestClassModule(projectDeclaration, "testSupertype2", true, null);
            var otherClass = GetTestClassModule(projectDeclaration, "otherTestClass", true, null);
            classModule.AddSupertype(supertype1);
            classModule.AddSupertype(supertype2);
            otherClass.AddSupertype(supertype1);
            otherClass.AddSupertype(supertype2);
            classModule.ClearSupertypes();

            Assert.IsFalse(supertype1.Subtypes.Any(subtype => subtype.Equals(classModule)));
            Assert.IsFalse(supertype2.Subtypes.Any(subtype => subtype.Equals(classModule)));
        }

        [Category("Resolver")]
        [Test]
        public void ClearSupertypeRemovesAllSupertypesDoesNotRemoveOtherSubtypesFromTheSupertypes()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var supertype1 = GetTestClassModule(projectDeclaration, "testSupertype1", true, null);
            var supertype2 = GetTestClassModule(projectDeclaration, "testSupertype2", true, null);
            var otherClass = GetTestClassModule(projectDeclaration, "otherTestClass", true, null);
            classModule.AddSupertype(supertype1);
            classModule.AddSupertype(supertype2);
            otherClass.AddSupertype(supertype1);
            otherClass.AddSupertype(supertype2);
            classModule.ClearSupertypes();

            Assert.IsTrue(supertype1.Subtypes.Any(subtype => subtype.Equals(otherClass)));
            Assert.IsTrue(supertype2.Subtypes.Any(subtype => subtype.Equals(otherClass)));
        }

        [Category("Resolver")]
        [Test]
        public void ByDefaultSupertypeNamesIsEmpty()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.SupertypeNames.Any());
        }


        [Category("Resolver")]
        [Test]
        public void AddSupertypeNameAddsTypenameToSupertypeNames()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var supertypeName = "testSupertypeName";
            classModule.AddSupertypeName(supertypeName);

            Assert.IsTrue(classModule.SupertypeNames.First().Equals(supertypeName));
        }


        [Category("Resolver")]
        [Test]
        public void AddSupertypeHasNoEffectOnSupertypeNames()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var supertype = GetTestClassModule(projectDeclaration, "testSupertype", true, null);
            classModule.AddSupertype(supertype);

            Assert.IsFalse(classModule.SupertypeNames.Any());
        }


        [Category("Resolver")]
        [Test]
        public void AddSupertypeNameHasNoEffectsOnSupertypes()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var supertypeName = "testSupertypeName";
            classModule.AddSupertypeName(supertypeName);

            Assert.IsFalse(classModule.Supertypes.Any());
        }


        [Category("Resolver")]
        [Test]
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


        [Category("Resolver")]
        [Test]
        public void GetSupertypesReturnsTheSupertypesOfAClassModule()
        {
            var projectDeclaration = GetTestProject("testProject");
            var supertype = GetTestClassModule(projectDeclaration, "testSupertype", true, null);
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            classModule.AddSupertype(supertype);

            Assert.AreEqual(supertype, ClassModuleDeclaration.GetSupertypes(classModule).First());
        }


        [Category("Resolver")]
        [Test]
        public void GetSupertypesReturnsAnEmptyEnumerableForDeclarationsWithDeclarationTypeClassModuleWhichAreNoClassModuleDeclarations()
        {
            var projectDeclaration = GetTestProject("testProject");
            var fakeClassModule = GetTestFakeClassModule(projectDeclaration, "testFakeClass");

            Assert.IsFalse(ClassModuleDeclaration.GetSupertypes(fakeClassModule).Any());
        }

            private static Declaration GetTestFakeClassModule(Declaration parentDeclatation, string name)
            {
                var qualifiedVariableMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new Declaration(qualifiedVariableMemberName, parentDeclatation, "dummy", "test", "test", false, false, Accessibility.Public, DeclarationType.ClassModule, null, null, Selection.Home, true, null);
            }



        [Category("Resolver")]
        [Test]
        public void ByDefaultDefaultMemberIsNull()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsNull(classModule.DefaultMember);
        }


        [Category("Resolver")]
        [Test]
        public void ByDefaultClassModulesNotBuiltInAreNotExposed()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.IsExposed);
        }


        // TODO: Find out if there's info about "being exposed" in type libraries.
        // We take the conservative approach of treating all type library modules as exposed.
        [Category("Resolver")]
        [Test]
        public void BuiltInClassesAreExposed()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", false, null);

            Assert.IsTrue(classModule.IsExposed);
        }


        [Category("Resolver")]
        [Test]
        public void ClassModulesWithTheExposedAttributeAreExposed()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddExposedClassAttribute();
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, classAttributes);

            Assert.IsTrue(classModule.IsExposed);
        }


        [Category("Resolver")]
        [Test]
        public void ByDefaultClassModulesAreNotGlobalClasses()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.IsGlobalClassModule);
        }


        [Category("Resolver")]
        [Test]
        public void ClassModulesWithTheGlobalNamespaceAttributeAreGlobalClasses()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddGlobalClassAttribute();
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, classAttributes);

            Assert.IsTrue(classModule.IsGlobalClassModule);
        }


        [Category("Resolver")]
        [Test]
        public void ClassModulesWithASubtypeBelowInTheHiearchyThatIsAGlobalClassAndThatHasBeenAddedBeforeCallingIsGlobalClassTheFirstTimeIsAGlobalClass()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddGlobalClassAttribute();
            var subsubtype = GetTestClassModule(projectDeclaration, "testSubSubtype", true, classAttributes);
            var subtype = GetTestClassModule(projectDeclaration, "testSubtype", true, null);
            subsubtype.AddSupertype(subtype);
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            subtype.AddSupertype(classModule);

            Assert.IsTrue(classModule.IsGlobalClassModule);
        }


        [Category("Resolver")]
        [Test]
        public void ClassModulesBecomeAGlobalClassIfASubtypeBelowInTheHiearchyIsAddedThatIsAGlobalClassAfterIsAGlobalClassHasAlreadyBeenCalled()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddGlobalClassAttribute();
            var subsubtype = GetTestClassModule(projectDeclaration, "testSubSubtype", true, classAttributes);
            var subtype = GetTestClassModule(projectDeclaration, "testSubtype", true, null);
            subsubtype.AddSupertype(subtype);
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            var dummy = classModule.IsGlobalClassModule;
            subtype.AddSupertype(classModule);

            Assert.IsTrue(classModule.IsGlobalClassModule);
        }


        [Category("Resolver")]
        [Test]
        public void ClassModulesBecomeAGlobalClassIfBelowInTheHierarchyASubtypeIsAddedThatIsAGlobalClassAfterIsAGlobalClassHasAlreadyBeenCalled()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddGlobalClassAttribute();
            var subsubtype = GetTestClassModule(projectDeclaration, "testSubSubtype", true, classAttributes);
            var subtype = GetTestClassModule(projectDeclaration, "testSubtype", true, null);
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);
            subtype.AddSupertype(classModule);
            var dummy = classModule.IsGlobalClassModule;
            subsubtype.AddSupertype(subtype);

            Assert.IsTrue(classModule.IsGlobalClassModule);
        }


        [Category("Resolver")]
        [Test]
        public void ByDefaultClassModulesDoNotHaveAPredeclaredID()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.HasPredeclaredId);
        }


        [Category("Resolver")]
        [Test]
        public void ClassModulesHaveAPredeclaredIDIfStatedInTheConstructorThatTheyHaveADefaultInstanceVariable()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null, true);

            Assert.IsTrue(classModule.HasPredeclaredId);
        }


        [Category("Resolver")]
        [Test]
        public void ClassModulesWithThePredeclaredIDAttributeHaveAPredeclaredID()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddPredeclaredIdTypeAttribute();
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, classAttributes);

            Assert.IsTrue(classModule.HasPredeclaredId);
        }


        [Category("Resolver")]
        [Test]
        public void ByDefaultClassModulesDoNotHaveADefaultInstanceVariable()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null);

            Assert.IsFalse(classModule.HasDefaultInstanceVariable);
        }


        [Category("Resolver")]
        [Test]
        public void ClassModulesThatAreGlobalClassesHaveADefaultInstanceVariable()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddGlobalClassAttribute();
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, classAttributes);

            Assert.IsTrue(classModule.HasDefaultInstanceVariable);
        }


        [Category("Resolver")]
        [Test]
        public void ClassModulesWithThePredeclaredIDAttributeHaveADefaultInstanceVariable()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classAttributes = new Attributes();
            classAttributes.AddPredeclaredIdTypeAttribute();
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, classAttributes);

            Assert.IsTrue(classModule.HasDefaultInstanceVariable);
        }


        [Category("Resolver")]
        [Test]
        public void ClassModulesHaveADefaultInstanceVariableIfThisIsStated()
        {
            var projectDeclaration = GetTestProject("testProject");
            var classModule = GetTestClassModule(projectDeclaration, "testClass", true, null, true);

            Assert.IsTrue(classModule.HasDefaultInstanceVariable);
        }

    }
}

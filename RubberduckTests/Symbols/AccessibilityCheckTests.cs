using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace RubberduckTests.Symbols
{
    [TestFixture]
    public class AccessibilityCheckTests
    {

        //project tests

        [Category("Resolver")]
        [Test]
        public void ProjectsAreAlwaysAccessible()
        {
            var projectDeclatation = GetTestProject("test_project");

            Assert.IsTrue(AccessibilityCheck.IsAccessible(null,null,null,projectDeclatation));
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



        //module tests

        [Category("Resolver")]
        [Test]
        public void ModulesCanBeAccessedFromWithinThemselves()
        {
            var projectDeclatation = GetTestProject("test_project");
            var moduleDeclatation = GetTestClassModule(projectDeclatation, "test_Module");

            Assert.IsTrue(AccessibilityCheck.IsAccessible(projectDeclatation, moduleDeclatation, null, moduleDeclatation));
        }

            private static ClassModuleDeclaration GetTestClassModule(Declaration projectDeclatation, string name, bool isExposed = false)
            {
                var qualifiedClassModuleMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                var classModuleAttributes = new Attributes();
                if (isExposed)
                {
                    classModuleAttributes.AddExposedClassAttribute();
                }
                return new ClassModuleDeclaration(qualifiedClassModuleMemberName, projectDeclatation, name, true, null, classModuleAttributes);
            }


        [Category("Resolver")]
        [Test]
        public void ModulesCanBeAccessedFromTheSameProject()
        {
            var projectDeclatation = GetTestProject("test_project");
            var calleeModuleDeclatation = GetTestClassModule(projectDeclatation, "calleeModule");
            var callingModuleDeclatation = GetTestClassModule(projectDeclatation, "callingModule");

            Assert.IsTrue(AccessibilityCheck.IsAccessible(projectDeclatation, callingModuleDeclatation, null, calleeModuleDeclatation));
        }


        [Category("Resolver")]
        [Test]
        public void ExposedClassModulesCanBeAccessedFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("calleeProject");
            var calleeModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "caleeModule", true);
            var callingProjectDeclatation = GetTestProject("callingProject");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "callingModule");

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, null, calleeModuleDeclatation));
        }


        [Category("Resolver")]
        [Test]
        public void NonExposedClassModulesCannotBeAccessedFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("calleeProject");
            var calleeModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "caleeModule", false);
            var callingProjectDeclatation = GetTestProject("callingProject");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "callingModule");

            Assert.IsFalse(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, null, calleeModuleDeclatation));
        }


        [Category("Resolver")]
        [Test]
        public void NonPrivateProceduralModulesCanBeAccessedFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("calleeProject");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "caleeModule");
            var callingProjectDeclatation = GetTestProject("callingProject");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "callingModule");

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, null, calleeModuleDeclatation));
        }

            private static ProceduralModuleDeclaration GetTestProceduralModule(Declaration projectDeclatation, string name)
            {
                var qualifiedClassModuleMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                var proceduralModuleDeclaration = new ProceduralModuleDeclaration(qualifiedClassModuleMemberName, projectDeclatation, name, true, null, null);
                return proceduralModuleDeclaration;
            }


        //TODO: Find a way to write PrivateProceduralModulesCannotBeAccessedFromOtherProjects. (isPrivateModule is a property with internal set.)



        //procedure tests

        [Category("Resolver")]
        [Test]
        public void PrivateProceduresAreAccessibleFromTheEnclosingModule()
        {
            var projectDeclatation = GetTestProject("test_project");
            var moduleDeclatation = GetTestClassModule(projectDeclatation, "test_Module");
            var privateCalleeFunctionDeclaration = GetTestFunction(moduleDeclatation, "calleeFoo", Accessibility.Private);
            var privateCallingFunctionDeclaration = GetTestFunction(moduleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(projectDeclatation, moduleDeclatation, privateCallingFunctionDeclaration, privateCalleeFunctionDeclaration));
        }

            private static FunctionDeclaration GetTestFunction(Declaration moduleDeclatation, string name, Accessibility functionAccessibility)
            {
                var qualifiedFunctionMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new FunctionDeclaration(qualifiedFunctionMemberName, moduleDeclatation, moduleDeclatation, "test", null, "test", functionAccessibility, null, null, Selection.Home, false, true, null, null);
            }

        [Category("Resolver")]
        [Test]
        public void PrivateProceduresAreAccessibleIfTheyAreInAClassAboveTheEnclosingModuleOfTheCallerInTheClassHierarchy()
        {
            var projectDeclatation = GetTestProject("test_project");
            var callingModule = GetTestClassModule(projectDeclatation, "callingModule");
            var privateCallingFunction = GetTestFunction(callingModule, "callingFoo", Accessibility.Private);
            var supertypeOfCallingModule = GetTestClassModule(projectDeclatation, "callingModuleSuper");
            callingModule.AddSupertype(supertypeOfCallingModule);
            var supertypeOfSupertypeOfCallingModule = GetTestClassModule(projectDeclatation, "callingModuleSuperSuper");
            supertypeOfCallingModule.AddSupertype(supertypeOfSupertypeOfCallingModule);
            var privateCalleeFunction = GetTestFunction(supertypeOfSupertypeOfCallingModule, "calleeFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(projectDeclatation, callingModule, privateCallingFunction, privateCalleeFunction));
        }


        [Category("Resolver")]
        [Test]
        public void PrivateProceduresAreNotAcessibleFromOtherUnrelatedModules()
        {
            var projectDeclatation = GetTestProject("test_project");
            var calleeModuleDeclatation = GetTestClassModule(projectDeclatation, "callee_test_Module");
            var calleeFunctionDeclaration = GetTestFunction(calleeModuleDeclatation, "calleeFoo", Accessibility.Private);
            var callingModuleDeclatation = GetTestClassModule(projectDeclatation, "calling_test_Module");
            var callingFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsFalse(AccessibilityCheck.IsAccessible(projectDeclatation, callingModuleDeclatation, callingFunctionDeclaration, calleeFunctionDeclaration));
        }


        [Category("Resolver")]
        [Test]
        public void FriendProceduresAreAcessibleFromOtherModulesInTheSameProject()
        {
            var projectDeclatation = GetTestProject("test_project");
            var calleeModuleDeclatation = GetTestClassModule(projectDeclatation, "callee_test_Module");
            var friendFunctionDeclaration = GetTestFunction(calleeModuleDeclatation, "calleeFoo", Accessibility.Friend);
            var callingModuleDeclatation = GetTestClassModule(projectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(projectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, friendFunctionDeclaration));
        }


        [Category("Resolver")]
        [Test]
        public void FriendProceduresAreNotAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var friendFunctionDeclaration = GetTestFunction(calleeModuleDeclatation, "calleeFoo", Accessibility.Friend);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsFalse(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, friendFunctionDeclaration));
        }


        [Category("Resolver")]
        [Test]
        public void PublicProceduresInExposedClassModulesAreAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var calleeFunctionDeclaration = GetTestFunction(calleeModuleDeclatation, "calleeFoo", Accessibility.Public);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, calleeFunctionDeclaration));
        }


        [Category("Resolver")]
        [Test]
        public void PublicProceduresInNonPrivateProceduralModulesAreAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var calleeFunctionDeclaration = GetTestFunction(calleeModuleDeclatation, "calleeFoo", Accessibility.Public);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, calleeFunctionDeclaration));
        }


        [Category("Resolver")]
        [Test]
        public void ImplicitelyScopedProceduresInExposedClassModulesAreAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var calleeFunctionDeclaration = GetTestFunction(calleeModuleDeclatation, "calleeFoo", Accessibility.Implicit);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, calleeFunctionDeclaration));
        }


        [Category("Resolver")]
        [Test]
        public void ImplicitelyScopedProceduresInNonPrivateProceduralModulesAreAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var calleeFunctionDeclaration = GetTestFunction(calleeModuleDeclatation, "calleeFoo", Accessibility.Implicit);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, calleeFunctionDeclaration));
        }



        //instance variable tests

        [Category("Resolver")]
        [Test]
        public void PrivateInstanceVariablesAreAccessibleFromTheEnclosingModule()     
        {
            var projectDeclatation = GetTestProject("test_project");
            var moduleDeclatation = GetTestClassModule(projectDeclatation, "test_Module");
            var privateFunctionDeclaration = GetTestFunction(moduleDeclatation, "foo", Accessibility.Private);
            var privateInstanceVariable = GetTestVariable(moduleDeclatation, "x", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(projectDeclatation,moduleDeclatation,privateFunctionDeclaration,privateInstanceVariable));
        }

            private static Declaration GetTestVariable(Declaration parentDeclatation, string name, Accessibility variableAccessibility)
            {
                var qualifiedVariableMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new Declaration(qualifiedVariableMemberName, parentDeclatation, "dummy", "test", "test", false, false, variableAccessibility, DeclarationType.Variable, null, null, Selection.Home, true, null);
            }


        [Category("Resolver")]
        [Test]
            public void PrivateInstanceVariablesAreAccessibleIfTheyAreInAClassAboveTheEnclosingModuleOfTheCallerInTheClassHierarchy()
        {
            var projectDeclatation = GetTestProject("test_project");
            var callingModule = GetTestClassModule(projectDeclatation, "callingModule");
            var privateCallingFunction = GetTestFunction(callingModule, "callingFoo", Accessibility.Private);
            var supertypeOfCallingModule = GetTestClassModule(projectDeclatation, "callingModuleSuper");
            callingModule.AddSupertype(supertypeOfCallingModule);
            var supertypeOfSupertypeOfCallingModule = GetTestClassModule(projectDeclatation, "callingModuleSuperSuper");
            supertypeOfCallingModule.AddSupertype(supertypeOfSupertypeOfCallingModule);
            var privateCalleeInstanceVariable = GetTestVariable(supertypeOfSupertypeOfCallingModule, "calleeFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(projectDeclatation, callingModule, privateCallingFunction, privateCalleeInstanceVariable));
        }


        [Category("Resolver")]
        [Test]
        public void PrivateInstanceVariablesAreNotAcessibleFromOtherUnrelatedModules()
        {
            var projectDeclatation = GetTestProject("test_project");
            var calleeModuleDeclatation = GetTestClassModule(projectDeclatation, "callee_test_Module");
            var instanceVariable = GetTestVariable(calleeModuleDeclatation, "x", Accessibility.Private);
            var callingModuleDeclatation = GetTestClassModule(projectDeclatation, "calliong_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "foo", Accessibility.Private);

            Assert.IsFalse(AccessibilityCheck.IsAccessible(projectDeclatation, callingModuleDeclatation, functionDeclaration, instanceVariable));
        }


        [Category("Resolver")]
        [Test]
        public void PublicInstanceVariablesInExposedClassModulesAreAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var instanceVariable = GetTestVariable(calleeModuleDeclatation, "x", Accessibility.Public);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, instanceVariable));
        }


        [Category("Resolver")]
        [Test]
        public void PublicInstanceVariablesInNonPrivateProceduralModulesAreAcessibleFromTheSameProject()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var instanceVariable = GetTestVariable(calleeModuleDeclatation, "x", Accessibility.Public);
            var callingModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(calleeProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, instanceVariable));
        }


        [Category("Resolver")]
        [Test] 
        public void PublicInstanceVariablesInNonPrivateProceduralModulesAreAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var instanceVariable = GetTestVariable(calleeModuleDeclatation, "x", Accessibility.Public);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, instanceVariable));
        }


        [Category("Resolver")]
        [Test] 
        public void GlobalInstanceVariablesInNonPrivateProceduralModulesAreAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var instanceVariable = GetTestVariable(calleeModuleDeclatation, "x", Accessibility.Global);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, instanceVariable));
        }


        [Category("Resolver")]
        [Test]
        public void ImplicitlyScopedInstanceVariablesAreAcessibleWithinTheirEnclosingModule()
        {
            var projectDeclatation = GetTestProject("test_project");
            var moduleDeclatation = GetTestClassModule(projectDeclatation, "test_Module");
            var privateFunctionDeclaration = GetTestFunction(moduleDeclatation, "foo", Accessibility.Private);
            var instanceVariable = GetTestVariable(moduleDeclatation, "x", Accessibility.Implicit);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(projectDeclatation, moduleDeclatation, privateFunctionDeclaration, instanceVariable));
        }


        [Category("Resolver")]
        [Test]
        public void ImplicitlyScopedInstanceVariablesAreNotAcessibleFromOtherModules()
        {
            var projectDeclatation = GetTestProject("test_project");
            var calleeModuleDeclatation = GetTestClassModule(projectDeclatation, "callee_test_Module");
            var instanceVariable = GetTestVariable(calleeModuleDeclatation, "x", Accessibility.Implicit);
            var callingModuleDeclatation = GetTestClassModule(projectDeclatation, "calliong_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "foo", Accessibility.Private);

            Assert.IsFalse(AccessibilityCheck.IsAccessible(projectDeclatation, callingModuleDeclatation, functionDeclaration, instanceVariable));
        }



        //local variable tests

        [Category("Resolver")]
        [Test]
        public void LocalMembersAreAcessibleFromTheMethodTheyAreDefinedIn()
        {
            var projectDeclatation = GetTestProject("test_project");
            var moduleDeclatation = GetTestClassModule(projectDeclatation, "test_Module");
            var privateFunctionDeclaration = GetTestFunction(moduleDeclatation, "foo", Accessibility.Private);
            var localVariable = GetTestVariable(privateFunctionDeclaration, "x", Accessibility.Implicit);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(projectDeclatation, moduleDeclatation, privateFunctionDeclaration, localVariable));
        }


        [Category("Resolver")]
        [Test]
        public void LocalMembersAreNotAcessibleFromOtherMethods()
        {
            var projectDeclatation = GetTestProject("test_project");
            var moduleDeclatation = GetTestClassModule(projectDeclatation, "test_Module");
            var privateFunctionDeclaration = GetTestFunction(moduleDeclatation, "calleeFoo", Accessibility.Private);
            var localVariable = GetTestVariable(privateFunctionDeclaration, "x", Accessibility.Implicit);
            var otherPrivateFunctionDeclaration = GetTestFunction(moduleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsFalse(AccessibilityCheck.IsAccessible(projectDeclatation, moduleDeclatation, otherPrivateFunctionDeclaration, localVariable));
        }


        [Category("Resolver")]
        [Test]
        public void StaticLocalMembersAreAcessibleFromTheMethodTheyAreDefinedIn()
        {
            var projectDeclatation = GetTestProject("test_project");
            var moduleDeclatation = GetTestClassModule(projectDeclatation, "test_Module");
            var privateFunctionDeclaration = GetTestFunction(moduleDeclatation, "foo", Accessibility.Private);
            var localVariable = GetTestVariable(privateFunctionDeclaration, "x", Accessibility.Static);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(projectDeclatation, moduleDeclatation, privateFunctionDeclaration, localVariable));
        }


        [Category("Resolver")]
        [Test]
        public void StaticLocalMembersAreNotAcessibleFromOtherMethods()
        {
            var projectDeclatation = GetTestProject("test_project");
            var moduleDeclatation = GetTestClassModule(projectDeclatation, "test_Module");
            var privateFunctionDeclaration = GetTestFunction(moduleDeclatation, "calleeFoo", Accessibility.Private);
            var localVariable = GetTestVariable(privateFunctionDeclaration, "x", Accessibility.Static);
            var otherPrivateFunctionDeclaration = GetTestFunction(moduleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsFalse(AccessibilityCheck.IsAccessible(projectDeclatation, moduleDeclatation, otherPrivateFunctionDeclaration, localVariable));
        }



        //enum tests

        [Category("Resolver")]
        [Test]
        public void PrivateEnumsAreAccessibleInTheSameModule()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var enumDeclarartion = GetTestEnum(calleeModuleDeclatation, "x", Accessibility.Private);
            var functionDeclaration = GetTestFunction(calleeModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(calleeProjectDeclatation, calleeModuleDeclatation, functionDeclaration, enumDeclarartion));
        }

            private static Declaration GetTestEnum(Declaration parentDeclatation, string name, Accessibility enumAccessibility)
            {
                var qualifiedVariableMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new Declaration(qualifiedVariableMemberName, parentDeclatation, "dummy", "test", "test", false, false, enumAccessibility, DeclarationType.Enumeration, null, null, Selection.Home, true, null);
            }


        [Category("Resolver")]
        [Test]
        public void PrivateEnumsAreNotAccessibleFromOtherModules()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var enumDeclarartion = GetTestEnum(calleeModuleDeclatation, "x", Accessibility.Private);
            var callingModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "calling_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsFalse(AccessibilityCheck.IsAccessible(calleeProjectDeclatation, callingModuleDeclatation, functionDeclaration, enumDeclarartion));
        }


        [Category("Resolver")]
        [Test]
        public void PublicEnumsInNonPrivateProceduralModulesAreAccessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var enumDeclarartion = GetTestEnum(calleeModuleDeclatation, "x", Accessibility.Public);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, functionDeclaration, enumDeclarartion));
        }


        [Category("Resolver")]
        [Test]
        public void GlobalEnumsInNonPrivateProceduralModulesAreAccessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var enumDeclarartion = GetTestEnum(calleeModuleDeclatation, "x", Accessibility.Global);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, functionDeclaration, enumDeclarartion));
        }


        [Category("Resolver")]
        [Test]
        public void ImplicitelyScopedEnumsInNonPrivateProceduralModulesAreAccessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var enumDeclarartion = GetTestEnum(calleeModuleDeclatation, "x", Accessibility.Implicit);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, functionDeclaration, enumDeclarartion));
        }


        [Category("Resolver")]
        [Test]
        public void PublicEnumsInExposedClassesAreNotAccessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclaration = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var enumDeclarartion = GetTestEnum(calleeModuleDeclaration, "x", Accessibility.Public);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, functionDeclaration, enumDeclarartion));
        }


        [Category("Resolver")]
        [Test]
        public void GlobalEnumsInExposedClassesAreAccessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclaration = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var enumDeclarartion = GetTestEnum(calleeModuleDeclaration, "x", Accessibility.Global);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, functionDeclaration, enumDeclarartion));
        }


        [Category("Resolver")]
        [Test]
        public void ImplicitelyScopedEnumsInExposedClassesAreAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclaration = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var enumDeclarartion = GetTestEnum(calleeModuleDeclaration, "x", Accessibility.Implicit);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, functionDeclaration, enumDeclarartion));
        }



        //user type tests


        [Category("Resolver")]
        [Test]
        public void PrivateUserTypesAreAccessibleInTheSameModule()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var userType = GetTestUserType(calleeModuleDeclatation, "x", Accessibility.Private);
            var functionDeclaration = GetTestFunction(calleeModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(calleeProjectDeclatation, calleeModuleDeclatation, functionDeclaration, userType));
        }

            private static Declaration GetTestUserType(Declaration parentDeclatation, string name, Accessibility userTypeAccessibility)
            {
                var qualifiedVariableMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new Declaration(qualifiedVariableMemberName, parentDeclatation, "dummy", "test", "test", false, false, userTypeAccessibility, DeclarationType.UserDefinedType, null, null, Selection.Home, true, null);
            }


        [Category("Resolver")]
        [Test]
        public void PrivateUserTypesAreNotAccessibleFromOtherModules()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var userType = GetTestUserType(calleeModuleDeclatation, "x", Accessibility.Private);
            var callingModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "calling_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsFalse(AccessibilityCheck.IsAccessible(calleeProjectDeclatation, callingModuleDeclatation, functionDeclaration, userType));
        }


        [Category("Resolver")]
        [Test]
        public void PublicUserTypesInNonPrivateProceduralModulesAreNotAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var userType = GetTestUserType(calleeModuleDeclatation, "x", Accessibility.Public);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, functionDeclaration, userType));
        }


        [Category("Resolver")]
        [Test]
        public void GlobalUserTypesInNonPrivateProceduralModulesAreNotAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var userType = GetTestUserType(calleeModuleDeclatation, "x", Accessibility.Global);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, functionDeclaration, userType));
        }


        [Category("Resolver")]
        [Test]
        public void ImplicitelyScopedUserTypesInNonPrivateProceduralModulesAreNotAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var userType = GetTestUserType(calleeModuleDeclatation, "x", Accessibility.Implicit);
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var functionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, functionDeclaration, userType));
        }


        //further tests (derived from code already present)

        [Category("Resolver")]
        [Test]
        public void EnumMembersInNonPrivateProceduralModulesAreNotAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var instanceVariable = GetTestEnumMember(calleeModuleDeclatation, "x");
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, instanceVariable));
        }

            private static Declaration GetTestEnumMember(Declaration parentDeclatation, string name)
            {
                var qualifiedVariableMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new Declaration(qualifiedVariableMemberName, parentDeclatation, "dummy", "test", "test", false, false, Accessibility.Implicit, DeclarationType.EnumerationMember, null, null, Selection.Home, true, null);
            }


        [Category("Resolver")]
        [Test]
        public void UserTypeMembersInExposedClassModulesAreAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestClassModule(calleeProjectDeclatation, "callee_test_Module", true);
            var instanceVariable = GetTestUserTypeMember(calleeModuleDeclatation, "x");
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, instanceVariable));
        }

            private static Declaration GetTestUserTypeMember(Declaration parentDeclatation, string name)
            {
                var qualifiedVariableMemberName = new QualifiedMemberName(StubQualifiedModuleName(), name);
                return new Declaration(qualifiedVariableMemberName, parentDeclatation, "dummy", "test", "test", false, false, Accessibility.Implicit, DeclarationType.UserDefinedTypeMember, null, null, Selection.Home, true, null);
            }


        [Category("Resolver")]
        [Test]
        public void UserTypeMembersInNonPrivateProceduralModulesAreNotAcessibleFromOtherProjects()
        {
            var calleeProjectDeclatation = GetTestProject("callee_test_project");
            var calleeModuleDeclatation = GetTestProceduralModule(calleeProjectDeclatation, "callee_test_Module");
            var instanceVariable = GetTestUserTypeMember(calleeModuleDeclatation, "x");
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsTrue(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, otherPrivateFunctionDeclaration, instanceVariable));
        }

        [Category("Resolver")]
        [Test]
        public void AccessibilityCheckDoesNotTakeIntoAccountThatAMemberMightNotBeAccessibleBecauseItIsCoveredByAnotherMemberInNarrowerScope()
        {
            var projectDeclatation = GetTestProject("test_project");
            var moduleDeclatation = GetTestClassModule(projectDeclatation, "test_Module");
            var privateInstanceVariable = GetTestVariable(moduleDeclatation, "x", Accessibility.Private);
            var privateFunctionDeclaration = GetTestFunction(moduleDeclatation, "foo", Accessibility.Private);
            var localVariable = GetTestVariable(privateFunctionDeclaration, "x", Accessibility.Implicit);   //This variable makes the instance variable of the same name inaccessible inside the function.

            Assert.IsTrue(AccessibilityCheck.IsAccessible(projectDeclatation, moduleDeclatation, privateFunctionDeclaration, privateInstanceVariable));
        }



        //null reference handling tests

        [Category("Resolver")]
        [Test]
        public void CalleesWhichAreNullAreNotAcessibleFromOtherProjects()
        {
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var callingFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsFalse(AccessibilityCheck.IsAccessible(callingProjectDeclatation, callingModuleDeclatation, callingFunctionDeclaration, null));
        }


        [Category("Resolver")]
        [Test]
        public void CalleeModulesWhichAreNullAreNotAcessibleFromOtherProjects()
        {
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var otherPrivateFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsFalse(AccessibilityCheck.IsModuleAccessible(callingProjectDeclatation, callingModuleDeclatation, null));
        }


        [Category("Resolver")]
        [Test]
        public void CalleeMembersWhichAreNullAreNotAcessibleFromOtherProjects()
        {
            var callingProjectDeclatation = GetTestProject("calling_test_project");
            var callingModuleDeclatation = GetTestClassModule(callingProjectDeclatation, "calling_test_Module");
            var callingFunctionDeclaration = GetTestFunction(callingModuleDeclatation, "callingFoo", Accessibility.Private);

            Assert.IsFalse(AccessibilityCheck.IsMemberAccessible(callingProjectDeclatation, callingModuleDeclatation, callingFunctionDeclaration, null));
        }


    }
}

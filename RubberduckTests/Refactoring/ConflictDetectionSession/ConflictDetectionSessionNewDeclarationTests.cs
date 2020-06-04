using NUnit.Framework;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Common;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System;
using System.Linq;
using TestResolver = RubberduckTests.Refactoring.ConflictDetectorTestsResolver;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ConflictDetectionSessionNewDeclarationTests
    {
        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Function)]
        [TestCase(DeclarationType.Procedure)]
        [TestCase(DeclarationType.Constant)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNameConflictsWithField(DeclarationType newDeclarationType)
        {
            var expectedName = "mTestVar1";
            var sourceCode =
$@"

Private Enum TestEnum
    AValue
End Enum

Private mTestVar As Long
";
            var nonConflictName = string.Empty;
            nonConflictName = RunNewDeclarationTestForNonConflictName(("mTestVar", newDeclarationType, Accessibility.Public), sourceCode);
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
        }

        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Function)]
        [TestCase(DeclarationType.Procedure)]
        [TestCase(DeclarationType.Constant)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNameConflictsWithFieldConstant(DeclarationType newDeclarationType)
        {
            var expectedName = "mTestVar1";
            var sourceCode =
$@"
Private Const MTestVAR As Long = 453
";
            var nonConflictName = RunNewDeclarationTestForNonConflictName(("mTestVar", newDeclarationType, Accessibility.Public), sourceCode);
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);

        }

        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Function)]
        [TestCase(DeclarationType.Procedure)]
        [TestCase(DeclarationType.Constant)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNameConflictsWithFieldFunction(DeclarationType newDeclarationType)
        {
            var expectedName = "Fizz1";
            var sourceCode =
$@"
Private Function Fizz() As Long
End Function
";
            var nonConflictName = RunNewDeclarationTestForNonConflictName(("Fizz", newDeclarationType, Accessibility.Public), sourceCode);
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
        }

        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Function)]
        [TestCase(DeclarationType.Procedure)]
        [TestCase(DeclarationType.Constant)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNameConflictsWithFieldSubroutine(DeclarationType newDeclarationType)
        {
            var expectedName = "Fazz1";
            var sourceCode =
$@"
Private Sub Fazz()
End Sub
";
            var nonConflictName = RunNewDeclarationTestForNonConflictName(("Fazz", newDeclarationType, Accessibility.Public), sourceCode);
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
        }


        [TestCase(DeclarationType.Variable, "Get Fazz() As Long")]
        [TestCase(DeclarationType.Variable, "Let Fazz(value As Long)")]
        [TestCase(DeclarationType.Variable, "Set Fazz(value As Long)")]
        [TestCase(DeclarationType.Function, "Get Fazz() As Long")]
        [TestCase(DeclarationType.Procedure, "Get Fazz() As Long")]
        [TestCase(DeclarationType.Constant, "Get Fazz() As Long")]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNameConflictsWithProperty(DeclarationType newDeclarationType, string signatureFragment)
        {
            var expectedName = "Fazz1";
            var sourceCode =
$@"
Private Property {signatureFragment}
End Property
";
            var nonConflictName = RunNewDeclarationTestForNonConflictName(("Fazz", newDeclarationType, Accessibility.Public), sourceCode);
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
        }

        [TestCase("Private Sub Fazz()\r\nEnd Sub")]
        [TestCase("Private Function Fazz() As Long\r\nEnd Function")]
        [TestCase("Private Property Get Fazz() As Long\r\nEnd Property")]
        [TestCase("Private Fazz As Long")]
        [TestCase("Private Const Fazz As Long = 6")]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewEnumMemberDeclarationNameConflicts(string declarationStatement)
        {
            var expectedName = "Fazz1";
            var sourceCode =
$@"
Public Enum TestEnum
    AValue
End Enum

{declarationStatement}
";

            var nonConflictName = RunNewDeclarationTestForNonConflictName(("Fazz", DeclarationType.EnumerationMember, Accessibility.Implicit), sourceCode, "TestEnum");
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
        }

        [TestCase(DeclarationType.Variable, MockVbeBuilder.TestModuleName)]
        [TestCase(DeclarationType.Function, MockVbeBuilder.TestModuleName)]
        [TestCase(DeclarationType.Procedure, MockVbeBuilder.TestModuleName)]
        [TestCase(DeclarationType.Constant, MockVbeBuilder.TestModuleName)]
        [TestCase(DeclarationType.EnumerationMember, "ETest")]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNameConflictsWithEnumMember(DeclarationType newDeclarationType, string parentDeclarationIdentifier)
        {
            var expectedName = "FirstValue1";
            var sourceCode =
$@"
Private Enum ETest
    FirstValue = 34
End Enum
";
            var nonConflictName = RunNewDeclarationTestForNonConflictName(("FirstValue", newDeclarationType, Accessibility.Public), sourceCode, parentDeclarationIdentifier);
            StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
        }


        [Test]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewDeclarationNamesRespectedByRename()
        {
            var sourceCode =
$@"
Private Sub Fizz()
End Sub
";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(sourceCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var expectedName = "SecondVariable1";

                var target = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Procedure)
                                .Where(d => d.IdentifierName.Equals("Fizz")).Single();

                var conflictSession = TestResolver.Resolve<IConflictSessionFactory>(state).Create();
                var moduleProxy = conflictSession.ProxyCreator.Create(target.QualifiedModuleName);
                foreach (var newVarName in new string[] {"FirstVariable", "SecondVariable" })
                {
                    var proxy = conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, DeclarationType.Variable, newVarName, Accessibility.Private);
                    conflictSession.TryRegister(proxy, out _, true);
                }


                var targetProxy = conflictSession.ProxyCreator.Create(target, "SecondVariable");
                conflictSession.TryRegister(targetProxy, out _, true);
                StringAssert.AreEqualIgnoringCase(expectedName, (conflictSession.RenamePairs.Single(pr => pr.Target == target)).NewName);
            }
        }

        [TestCase(MockVbeBuilder.TestModuleName)]
        [TestCase(MockVbeBuilder.TestProjectName)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewModuleDeclarations(string newModuleName)
        {
            var sourceCode =
$@"
Private mTest As Long
";
            CreateAndParseAndTest(sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var expectedName = newModuleName == MockVbeBuilder.TestModuleName
                        ? IncrementIdentifier(MockVbeBuilder.TestModuleName)
                        : IncrementIdentifier(MockVbeBuilder.TestProjectName);

                var newModuleProxy = testParams.conflictSession.ProxyCreator.CreateNewModule(testParams.stdModule.ProjectId, ComponentType.StandardModule, newModuleName);
                testParams.conflictSession.NewModuleConflictDetector
                                                .HasConflictingName(newModuleProxy, out var nonConflictName);
                StringAssert.AreEqualIgnoringCase(expectedName, nonConflictName);
            }
        }

        [TestCase(DeclarationType.Enumeration, false)]
        [TestCase(DeclarationType.Enumeration, true)]
        [TestCase(DeclarationType.UserDefinedType, false)]
        [TestCase(DeclarationType.UserDefinedType, true)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewModuleWithEnumOrUDTDeclaration(DeclarationType declarationType, bool forceRegistration)
        {
            var sourceCode =
$@"
Private mTest As Long
";
            CreateAndParseAndTest(sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var moduleProxy = testParams.conflictSession.ProxyCreator.CreateNewModule(testParams.stdModule.ProjectId, ComponentType.StandardModule, "SomethingNew");
                testParams.conflictSession.TryRegister(moduleProxy, out _, true);

                var proxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, declarationType, "SomethingNew", Accessibility.Public);

                var registerSuccess = testParams.conflictSession.TryRegister(proxy, out var nonConflictName, forceRegistration);

                Assert.AreEqual(forceRegistration, registerSuccess);

                StringAssert.AreEqualIgnoringCase("SomethingNew1", nonConflictName);
            }
        }

        [TestCase(DeclarationType.UserDefinedType, DeclarationType.Enumeration)]
        [TestCase(DeclarationType.Procedure, DeclarationType.Function)]
        [TestCase(DeclarationType.Variable, DeclarationType.Constant)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewModuleLevelDeclarationsWithSameName(DeclarationType first, DeclarationType second)
        {
            var sourceCode =
$@"
Private mTest As Long
";
            CreateAndParseAndTest(sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var initialIdentifier = "SomethingNew";

                var moduleProxy = testParams.conflictSession.ProxyCreator.Create(testParams.stdModule.QualifiedModuleName);
                var firstProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, first, initialIdentifier, Accessibility.Private);
                testParams.conflictSession.TryRegister(firstProxy, out _, true);

                var secondProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, second, initialIdentifier, Accessibility.Private);
                testParams.conflictSession.TryRegister(secondProxy, out _, true);

                StringAssert.AreEqualIgnoringCase("SomethingNew1", secondProxy.IdentifierName);
            }
        }

        [TestCase(DeclarationType.UserDefinedType)]
        [TestCase(DeclarationType.Enumeration)]
        [TestCase(DeclarationType.Procedure)]
        [TestCase(DeclarationType.Variable)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void SequenceOfDeclarationsWithSameName(DeclarationType declarationType)
        {
            var sourceCode =
$@"
Private mTest As Long
";
            CreateAndParseAndTest(sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var identifier = "SameName";

                var moduleProxy = testParams.conflictSession.ProxyCreator.Create(testParams.stdModule.QualifiedModuleName);

                var idx = 0;
                for (idx = 0; idx < 5; idx++)
                {
                    var proxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, declarationType, identifier, Accessibility.Private);
                    testParams.conflictSession.TryRegister(proxy, out _, true);
                }

                Assert.AreEqual(5, testParams.conflictSession.RegisteredProxies.Count);

                for (idx = 0; idx < testParams.conflictSession.RegisteredProxies.Count; idx++)
                {
                    var expectedName = idx > 0 ? $"{identifier}{idx}" : identifier;
                    var proxy = testParams.conflictSession.RegisteredProxies.SingleOrDefault(pr => pr.IdentifierName.Equals(expectedName));
                    StringAssert.AreEqualIgnoringCase(expectedName, proxy?.IdentifierName);
                }
            }
        }

        [TestCase(DeclarationType.Variable)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewModuleWithSequenceOfDeclarationsWithSameName(DeclarationType declarationType)
        {
            var sourceCode =
$@"
Private mTest As Long
";
            CreateAndParseAndTest(sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var identifier = "SameName";
                var moduleProxy = testParams.conflictSession.ProxyCreator.CreateNewModule(testParams.stdModule.ProjectId, ComponentType.StandardModule, "NewModule");
                testParams.conflictSession.TryRegister(moduleProxy, out _, true);
                var idx = 0;
                for (idx = 0; idx < 5; idx++)
                {
                    var proxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, declarationType, identifier);
                    testParams.conflictSession.TryRegister(proxy, out _, true);
                }

                Assert.AreEqual(6, testParams.conflictSession.RegisteredProxies.Count);

                for (idx = 0; idx < testParams.conflictSession.RegisteredProxies.Count - 1; idx++)
                {
                    var expectedName = idx > 0 ? $"{identifier}{idx}" : identifier;
                    var proxy = testParams.conflictSession.RegisteredProxies
                        .SingleOrDefault(pr => pr.IdentifierName.Equals(expectedName));
                    StringAssert.AreEqualIgnoringCase(expectedName, proxy?.IdentifierName);
                }
            }
        }

        [Test]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void SequenceOfModuleDeclarationsWithSameName()
        {
            var sourceCode =
$@"
Private mTest As Long
";
            CreateAndParseAndTest(sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var identifier = "SameName";
                var idx = 0;
                for (idx = 0; idx < 5; idx++)
                {
                    var proxy = testParams.conflictSession.ProxyCreator.CreateNewModule(testParams.stdModule.ProjectId, ComponentType.StandardModule, identifier);
                    testParams.conflictSession.TryRegister(proxy, out _, true);
                }

                Assert.AreEqual(5, testParams.conflictSession.RegisteredProxies.Count);

                for (idx = 0; idx < testParams.conflictSession.RegisteredProxies.Count; idx++)
                {
                    var expectedName = idx > 0 ? $"{identifier}{idx}" : identifier;
                    var proxy = testParams.conflictSession.RegisteredProxies.SingleOrDefault(pr => pr.IdentifierName.Equals(expectedName));
                    Assert.IsNotNull(proxy);
                }
            }
        }

        [TestCase("arg", "arg1")]
        [TestCase("fizz", "fizz1")]
        [TestCase("localVar", "localVar1")]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewLocalVariable(string newDeclarationIdentifier, string expected)
        {
            var sourceCode =
$@"
Private mTest As Long

Private Function Fizz(arg As Long) As String
    Dim localVar As String
    localVar = ""Yo""
    Fizz = localVar
End Function
";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(sourceCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var stdModule = state.DeclarationFinder.MatchName(MockVbeBuilder.TestModuleName).Cast<ModuleDeclaration>().Single();

                var procedure = state.DeclarationFinder.MatchName("Fizz").Single();

                var conflictSession = TestResolver.Resolve<IConflictSessionFactory>(state).Create();
                var parentProxy = conflictSession.ProxyCreator.Create(procedure);
                var proxy = conflictSession.ProxyCreator.CreateNewEntity(parentProxy, DeclarationType.Variable, newDeclarationIdentifier);
                conflictSession.TryRegister(proxy, out _, true);

                Assert.IsTrue(conflictSession.RegisteredProxies.Any());

                StringAssert.AreEqualIgnoringCase(expected, proxy.IdentifierName);
            }
        }


        [TestCase(DeclarationType.Variable)]
        [TestCase(DeclarationType.Constant)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewModuleAndNonMember_RenamesConflictingNewMember(DeclarationType nonMemberType)
        {
            var sourceCode =
$@"
Option Explicit
";
            CreateAndParseAndTest(sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var moduleProxy = testParams.conflictSession.ProxyCreator.CreateNewModule(testParams.stdModule.ProjectId, ComponentType.StandardModule, "Fizz");
                testParams.conflictSession.TryRegister(moduleProxy, out _);

                var nonMemberProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, nonMemberType, "fizz", Accessibility.Private);
                testParams.conflictSession.TryRegister(nonMemberProxy, out _, true);

                var memberProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, DeclarationType.Function, "Fizz", Accessibility.Public);
                testParams.conflictSession.TryRegister(memberProxy, out _, true);

                Assert.IsTrue(testParams.conflictSession.RegisteredProxies.Any());

                StringAssert.AreEqualIgnoringCase("Fizz", moduleProxy.IdentifierName);
                StringAssert.AreEqualIgnoringCase("fizz", nonMemberProxy.IdentifierName);
                StringAssert.AreEqualIgnoringCase("Fizz1", memberProxy.IdentifierName);
            }
        }

        [TestCase("blah", "blah1")]
        [TestCase("nextArg", "nextArg1")]
        [TestCase("local", "local1")]
        [TestCase("fizz", "fizz")]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewParameters(string newName, string expectedName)
        {
            var sourceCode =
$@"
Private Function Blah(arg As Long, nextArg As String) As Long
    Dim local As Long
    local = 6
    Blah = local + arg
End Function

";
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(sourceCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var function = state.DeclarationFinder.DeclarationsWithType(DeclarationType.Function).Single();

                var conflictSession = TestResolver.Resolve<IConflictSessionFactory>(state).Create();
                var parentProxy = conflictSession.ProxyCreator.Create(function);
                var parameterProxy = conflictSession.ProxyCreator.CreateNewEntity(parentProxy, DeclarationType.Parameter, newName);
                conflictSession.TryRegister(parameterProxy, out _, true);

                StringAssert.AreEqualIgnoringCase(expectedName, parameterProxy.IdentifierName);
            }
        }

        [TestCase(DeclarationType.Function)]
        [TestCase(DeclarationType.Procedure)]
        [TestCase(DeclarationType.PropertyLet)]
        [TestCase(DeclarationType.PropertySet)]
        [TestCase(DeclarationType.PropertyGet)]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewMemberWithParameters(DeclarationType declarationType)
        {
            var sourceCode =
$@"
Option Explicit

Private mTest As Long

";
            CreateAndParseAndTest(sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var argumentIdentifier = "arg";

                var moduleProxy = testParams.conflictSession.ProxyCreator.Create(testParams.stdModule.QualifiedModuleName);
                var memberProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, declarationType, "NewMember", Accessibility.Public);
                testParams.conflictSession.TryRegister(memberProxy, out _);

                var idx = 0;
                for (idx = 0; idx < 5; idx++)
                {
                    var parameterProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(memberProxy, DeclarationType.Parameter, argumentIdentifier);
                    testParams.conflictSession.TryRegister(parameterProxy, out _, true);
                }

                for (idx = 0; idx < 5; idx++)
                {
                    var expectedName = idx > 0 ? $"{argumentIdentifier}{idx}" : argumentIdentifier;
                    var proxy = testParams.conflictSession.RegisteredProxies.FirstOrDefault(pr => pr.IdentifierName.Equals($"{expectedName}"));
                    StringAssert.AreEqualIgnoringCase(expectedName, proxy?.IdentifierName);
                }
            }
        }

        [Test]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void EnumerationAndMembers()
        {
            var sourceCode =
$@"
Option Explicit

Private TestEnum3 As Long

";
            CreateAndParseAndTest(sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var moduleProxy = testParams.conflictSession.ProxyCreator.Create(testParams.stdModule.QualifiedModuleName);
                var enumerationProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, DeclarationType.Enumeration, "MyTestEnums", Accessibility.Public);
                testParams.conflictSession.TryRegister(enumerationProxy, out _);

                for (var idx = 0; idx < 5; idx++)
                {
                    var enumMemberProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(enumerationProxy, DeclarationType.EnumerationMember, "TestEnum");
                    testParams.conflictSession.TryRegister(enumMemberProxy, out _, true);
                }

                var expectedName = "TestEnum5";
                var proxy = testParams.conflictSession.RegisteredProxies.FirstOrDefault(pr => pr.IdentifierName.Equals(expectedName));
                StringAssert.AreEqualIgnoringCase(expectedName, proxy?.IdentifierName);

                expectedName = "TestEnum3"; //No proxy will have this identifier
                proxy = testParams.conflictSession.RegisteredProxies.FirstOrDefault(pr => pr.IdentifierName.Equals(expectedName));
                Assert.IsNull(proxy);
            }
        }

        [TestCase("MyTestUDT", "MyTestUDT")]
        [TestCase("MyTestEnum", "myTestEnum1")]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewUDT(string udtName, string expectedName)
        {
            var sourceCode =
$@"
Option Explicit

Public Enum MyTestEnum
    Blah
    BlahBlah
End Enum

";
            CreateAndParseAndTest(sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var moduleProxy = testParams.conflictSession.ProxyCreator.Create(testParams.stdModule.QualifiedModuleName);

                var udtProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, DeclarationType.UserDefinedType, udtName, Accessibility.Public);
                testParams.conflictSession.TryRegister(udtProxy, out _, true);

                for (var idx = 0; idx < 5; idx++)
                {
                    var udtMemberProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(udtProxy, DeclarationType.UserDefinedTypeMember, "MyTestUDTMember");
                    testParams.conflictSession.TryRegister(udtMemberProxy, out _, true);
                }

                var proxy = testParams.conflictSession.RegisteredProxies.FirstOrDefault(pr => pr.IdentifierName.Equals(expectedName, StringComparison.InvariantCultureIgnoreCase));
                StringAssert.AreEqualIgnoringCase(expectedName, proxy?.IdentifierName);

                expectedName = "MyTestUDTMember3";
                proxy = testParams.conflictSession.RegisteredProxies.FirstOrDefault(pr => pr.IdentifierName.Equals(expectedName));
                StringAssert.AreEqualIgnoringCase(expectedName, proxy?.IdentifierName);
            }
        }

        [Test]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewUDTMembers()
        {
            var sourceCode =
$@"
Option Explicit

Private testUDTMember3 As Long

";
            CreateAndParseAndTest(sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var moduleProxy = testParams.conflictSession.ProxyCreator.Create(testParams.stdModule.QualifiedModuleName);
                var udtProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(moduleProxy, DeclarationType.UserDefinedType, "MyTestUDT", Accessibility.Public);
                testParams.conflictSession.TryRegister(udtProxy, out _);
                var udtMemberConflictDetector = testParams.conflictSession.NewEntityConflictDetector;
                var memberName = "MyTestUDTMember";
                for (var idx = 0; idx < 5; idx++)
                {
                    var udtMemberProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(udtProxy, DeclarationType.UserDefinedTypeMember, memberName);
                    testParams.conflictSession.TryRegister(udtMemberProxy, out _, true);
                }

                var expectedName = $"{memberName}4";
                var proxy = testParams.conflictSession.RegisteredProxies.FirstOrDefault(pr => pr.IdentifierName.Equals(expectedName));
                StringAssert.AreEqualIgnoringCase(expectedName, proxy?.IdentifierName);

                expectedName = $"{memberName}5"; //No proxy will have this identifier
                proxy = testParams.conflictSession.RegisteredProxies.FirstOrDefault(pr => pr.IdentifierName.Equals(expectedName));
                Assert.IsNull(proxy);
            }
        }


        [TestCase(ComponentType.StandardModule, DeclarationType.PropertyGet, "testEntity1")]
        [TestCase(ComponentType.StandardModule, DeclarationType.Function, "testEntity1")]
        [TestCase(ComponentType.ClassModule, DeclarationType.PropertyGet, "testEntity")]
        [TestCase(ComponentType.ClassModule, DeclarationType.Function, "testEntity")]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewModuleWithPublicMembers(ComponentType newModuleComponentType, DeclarationType newEntityDeclarationType, string expectedNonConflictName)
        {
            var conflictName = "testEntity";
            var sourceCode =
$@"
Option Explicit

Public {conflictName} As Long

";
            var referenceSourceCode =
$@"
Option Explicit

Public Function DoIt() As Long
    DoIt = {conflictName}
End Function
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, sourceCode), ("ReferencingModule", referenceSourceCode));
            CreateAndParseAndTest(vbe.Object, sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var newModuleProxy = testParams.conflictSession.ProxyCreator.CreateNewModule(testParams.stdModule.ProjectId, newModuleComponentType, "NewModule");
                var newPublicEntityProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(newModuleProxy, newEntityDeclarationType, conflictName, Accessibility.Public);

                var result = testParams.conflictSession.NewEntityConflictDetector.HasConflictingName(newPublicEntityProxy, out var nonConflictName);
                StringAssert.AreEqualIgnoringCase(expectedNonConflictName, nonConflictName);
            }
        }


        [TestCase(ComponentType.StandardModule, DeclarationType.Variable, "testEntity1")]
        [TestCase(ComponentType.StandardModule, DeclarationType.Constant, "testEntity1")]
        [TestCase(ComponentType.ClassModule, DeclarationType.Variable, "testEntity")]
        [TestCase(ComponentType.ClassModule, DeclarationType.Constant, "testEntity")]
        [Category("Refactoring")]
        [Category(nameof(ConflictDetectionSession))]
        public void NewModuleWithPublicNonMembers(ComponentType newModuleComponentType, DeclarationType newEntityDeclarationType, string expectedNonConflictName)
        {
            var conflictName = "testEntity";
            var sourceCode =
$@"
Option Explicit

Public {conflictName} As Long

";
            var referenceSourceCode =
$@"
Option Explicit

Public Function DoIt() As Long
    DoIt = {conflictName}
End Function
";
            var vbe = MockVbeBuilder.BuildFromStdModules((MockVbeBuilder.TestModuleName, sourceCode), ("ReferencingModule", referenceSourceCode));
            CreateAndParseAndTest(vbe.Object, sourceCode, ThisTest);

            void ThisTest((ModuleDeclaration stdModule, IConflictSession conflictSession) testParams)
            {
                var newModuleProxy = testParams.conflictSession.ProxyCreator.CreateNewModule(testParams.stdModule.ProjectId, newModuleComponentType, "NewModule");
                var newPublicEntityProxy = testParams.conflictSession.ProxyCreator.CreateNewEntity(newModuleProxy, newEntityDeclarationType, conflictName, Accessibility.Public);

                var result = testParams.conflictSession.NewEntityConflictDetector.HasConflictingName(newPublicEntityProxy, out var nonConflictName);
                StringAssert.AreEqualIgnoringCase(expectedNonConflictName, nonConflictName);
            }
        }

        private static string RunNewDeclarationTestForNonConflictName((string ID, DeclarationType Type, Accessibility accessibility) target, string sourceCode, string parentDeclarationIdentifier = null)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(sourceCode, out _);
            var state = MockParser.CreateAndParse(vbe.Object);
            using (state)
            {
                var stdModule = state.DeclarationFinder.MatchName(MockVbeBuilder.TestModuleName).Cast<ModuleDeclaration>().Single();

                var parentDeclaration = string.IsNullOrEmpty(parentDeclarationIdentifier)
                        ? stdModule
                        : state.DeclarationFinder.MatchName(parentDeclarationIdentifier).Single();

                var session = TestResolver.Resolve<IConflictSessionFactory>(state).Create();
                IConflictDetectionDeclarationProxy testProxy = null;
                if (parentDeclaration.DeclarationType.HasFlag(DeclarationType.Module))
                {
                    var moduleProxy = session.ProxyCreator.Create(stdModule.QualifiedModuleName);

                    var proxy = session.ProxyCreator.CreateNewEntity(moduleProxy, target.Type, target.ID, target.accessibility);
                    session.TryRegister(proxy, out _, true);
                    testProxy = proxy;
                }
                else
                {
                    var parentProxy = session.ProxyCreator.Create(parentDeclaration);
                    var proxy = session.ProxyCreator.CreateNewEntity(parentProxy, target.Type, target.ID);
                    session.TryRegister(proxy, out _, true);
                    testProxy = proxy;
                }

                var newDeclarationProxy = session.RegisteredProxies.Where(p => p == testProxy).SingleOrDefault();
                return newDeclarationProxy?.IdentifierName ?? string.Empty;
            }
        }

        private static void CreateAndParseAndTest(string sourceCode, Action<(ModuleDeclaration, IConflictSession)> theTest)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(sourceCode, out _);
            CreateAndParseAndTest(vbe.Object, sourceCode, theTest);
        }

        private static void CreateAndParseAndTest(IVBE vbe, string sourceCode, Action<(ModuleDeclaration, IConflictSession)> theTest)
        {
            var state = MockParser.CreateAndParse(vbe);
            using (state)
            {
                var stdModule = state.DeclarationFinder.MatchName(MockVbeBuilder.TestModuleName).Cast<ModuleDeclaration>().Single();
                var conflictSession = TestResolver.Resolve<IConflictSessionFactory>(state).Create();
                theTest((stdModule, conflictSession));
            }
        }

        private static string IncrementIdentifier(string identifier)
        {
            var numeric = string.Concat(identifier.Reverse().TakeWhile(c => char.IsDigit(c)).Reverse());
            if (!int.TryParse(numeric, out var currentNum))
            {
                currentNum = 0;
            }
            var identifierSansNumericSuffix = identifier.Substring(0, identifier.Length - numeric.Length);
            return $"{identifierSansNumericSuffix}{++currentNum}";
        }
    }
}

using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;
using Antlr4.Runtime;
using Rubberduck.Inspections;

namespace RubberduckTests.Symbols
{
    [TestClass]
    public class DeclarationFinderTests
    {
        private IVBE _accessibilityTests_VBE;
        private ParseCoordinator _accessibilityTests_parser;
        private IEnumerable<Declaration> _accessibilityTests_declarations;
        private List<TestComponentSpecification> _accessibilityTests_Components;
        private List<string> _accessibilityTests_ProcedureScopeNames;
        private List<string> _accessibilityTests_ModuleScopeNames;
        private List<string> _accessibilityTests_GlobalScopeNames;

        [TestMethod]
        [TestCategory("Resolver")]
        public void DeclarationFinder_FindsAccessibleDeclarations_InProcedure()
        {
            SetupSUT("TestProject");
            TestAccessibleDeclarations(_accessibilityTests_ProcedureScopeNames, "ProcedureScope", true, false);
            
        }

        [TestMethod]
        [TestCategory("Resolver")]
        public void DeclarationFinder_FindsAccessibleDeclarations_ModuleScope()
        {
            SetupSUT("TestProject");
            TestAccessibleDeclarations(_accessibilityTests_ModuleScopeNames, "ModuleScope", true, false);
        }

        [TestMethod]
        [TestCategory("Resolver")]
        public void DeclarationFinder_FindsAccessibleDeclarations_GlobalScope()
        {
            SetupSUT("TestProject");
            TestAccessibleDeclarations(_accessibilityTests_GlobalScopeNames, "GlobalScope", true, true);
        }

        [TestMethod]

        [TestCategory("Resolver")]
        public void DeclarationFinder_FindsAccessibleDeclarations_Inaccessible()
        {
            SetupSUT("TestProject");
            string[] inaccessible = {"result", "mySecondEggo","localVar" , "FooFighters" , "filename", "implicitVar"};
            
            TestAccessibleDeclarations(inaccessible.ToList(), "GlobalScope", false, false);
        }

        [TestMethod]
        [TestCategory("Resolver")]
        public void DeclarationFinder_FindsAccessibleDeclarations_All()
        {
            //Tests that the DeclarationFinder does not return any unexpected declarations 
            //excluding Reserved Identifiers
            SetupSUT("TestProject");
            var allNamesUsedInTests = _accessibilityTests_ProcedureScopeNames;
            allNamesUsedInTests.AddRange(_accessibilityTests_ModuleScopeNames);
            allNamesUsedInTests.AddRange(_accessibilityTests_GlobalScopeNames);
            allNamesUsedInTests.AddRange(_accessibilityTests_Components.Select(n => n.Name));

            var target = GetTargetForAccessibilityTests();

            var declarationFinderResults =
                _accessibilityTests_parser.State.DeclarationFinder.GetDeclarationsAccessibleToScope(target, _accessibilityTests_declarations);
            var accessibleNames = declarationFinderResults.Select(d => d.IdentifierName);

            var unexpectedDeclarations = accessibleNames.Except(allNamesUsedInTests).ToList().Where(n => !VariableNameValidator.IsReservedIdentifier(n));

            string failureMessage = string.Empty;
            if( unexpectedDeclarations.Count() > 0)
            {
                failureMessage = unexpectedDeclarations.Count().ToString() + " unexpected declaration(s) found:";
                foreach(string identifier in unexpectedDeclarations)
                {
                    failureMessage = failureMessage + " '" + identifier + "', ";
                }
                failureMessage = failureMessage.Substring(0, failureMessage.Length - 2);
            }

            Assert.AreEqual(0, unexpectedDeclarations.Count(), failureMessage);
        }

        private void TestAccessibleDeclarations(IEnumerable<string> namesToTest, string scope, bool containsIdentifiers, bool includeModuleNames)
        {

            var allIdentifiersToCheck = namesToTest.ToList();
            if (includeModuleNames)
            {
                allIdentifiersToCheck.AddRange(_accessibilityTests_Components.Select(n => n.Name));
            }

            var target = GetTargetForAccessibilityTests();

            var declarationFinderResults = 
                _accessibilityTests_parser.State.DeclarationFinder.GetDeclarationsAccessibleToScope( target, _accessibilityTests_declarations);

            var accessibleNames = declarationFinderResults.Select(d => d.IdentifierName);

            var messagePreface = "Test failed for  " + scope + " identifier: ";
            foreach (var identifier in allIdentifiersToCheck)
            {
                if (containsIdentifiers)
                {
                    Assert.IsTrue(accessibleNames.Contains(identifier), messagePreface + identifier);
                }
                else
                {
                    Assert.IsFalse(accessibleNames.Contains(identifier), messagePreface + identifier);
                }
            }
        }
        private Declaration GetTargetForAccessibilityTests()
        {
            var targetIdentifier = "targetAccessibilityTests";
            var targets = _accessibilityTests_declarations.Where(dec => dec.IdentifierName == targetIdentifier);
            if (targets.Count() > 1)
            {
                Assert.Inconclusive("Multiple targets found with identifier: " + targetIdentifier + ".  Test requires a unique identifierName");
            }
            return targets.FirstOrDefault();
        }

        private void SetupSUT(string projectName)
        {
            if(_accessibilityTests_VBE != null ) { return; }

            string[] accessibleWithinParentProcedure = { "arg1", "FooBar1", "targetAccessibilityTests", "theSecondArg" };
            string[] accessibleModuleScope = { "memberString", "memberLong", "myEggo", "Foo", "FooBar1", "GoMyEggo", "FooFight" };
            string[] accessibleGlobalScope = { "CantTouchThis", "BigNumber", "DoSomething", "SetFilename", "ShortStory","THE_FILENAME"};

            _accessibilityTests_ProcedureScopeNames = accessibleWithinParentProcedure.ToList();
            _accessibilityTests_ModuleScopeNames = accessibleModuleScope.ToList();
            _accessibilityTests_GlobalScopeNames = accessibleGlobalScope.ToList();

            var firstClassBody = GetRespectsDeclarationsAccessibilityRules_FirstClassBody();
            var secondClassBody = GetRespectsDeclarationsAccessibilityRules_SecondClassBody();
            var firstModuleBody = GetRespectsDeclarationsAccessibilityRules_FirstModuleBody();
            var secondModuleBody = GetRespectsDeclarationsAccessibilityRules_SecondModuleBody();

            _accessibilityTests_Components = new List<TestComponentSpecification>();
            _accessibilityTests_Components.Add( new TestComponentSpecification("CFirstClass", firstClassBody, ComponentType.ClassModule));
            _accessibilityTests_Components.Add(new TestComponentSpecification("CSecondClass", secondClassBody, ComponentType.ClassModule));
            _accessibilityTests_Components.Add(new TestComponentSpecification("modFirst", firstModuleBody, ComponentType.StandardModule));
            _accessibilityTests_Components.Add(new TestComponentSpecification("modSecond", secondModuleBody, ComponentType.StandardModule));


            _accessibilityTests_VBE = BuildProject("TestProject", _accessibilityTests_Components);

            _accessibilityTests_parser = MockParser.Create(_accessibilityTests_VBE, new RubberduckParserState(_accessibilityTests_VBE));
            _accessibilityTests_parser.Parse(new CancellationTokenSource());
            if (_accessibilityTests_parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            _accessibilityTests_declarations = _accessibilityTests_parser.State.AllUserDeclarations;
        }

        private IVBE BuildProject(string projectName, List<TestComponentSpecification> testComponents)
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected);

            testComponents.ForEach(c => enclosingProjectBuilder.AddComponent(c.Name, c.ModuleType, c.Content));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            return builder.Build().Object;
        }

        internal class TestComponentSpecification
        {
            private string _name;
            private string _content;
            private ComponentType _componentType;
            public TestComponentSpecification(string componentName, string componentContent, ComponentType componentType)
            {
                _name = componentName;
                _content = componentContent;
                _componentType = componentType;
            }

            public string Name { get { return _name; } }
            public string Content { get { return _content; } }
            public ComponentType ModuleType { get { return _componentType; } }
        }

#region AccessibilityTestsModuleContent
        private string GetRespectsDeclarationsAccessibilityRules_FirstClassBody()
        {
            return
@"
Private memberString As String
Private memberLong As Long
Private myEggo As String

Public Sub Foo(ByVal arg1 As String)
    Dim localVar as Long
    localVar = 7
    Let arg1 = ""test""
    memberString = arg1 & ""Foo""
End Sub

Public Function FooBar1(ByRef arg1 As String, theSecondArg As Long) As String
    Let arg1 = ""test""
    Dim targetAccessibilityTests As String
    targetAccessibilityTests = arg1 & CStr(theSecondArg)
    FooBar1 = targetAccessibilityTests
End Function

Property Let GoMyEggo(newValue As String)
    myEggo = newValue
End Property

Property Get GoMyEggo()
    GoMyEggo = myEggo
End Property

Private Sub FooFight(ByRef arg1 As String)
    xArg1 = 6
    Let arg1 = ""test""
End Sub
";
        }

        private string GetRespectsDeclarationsAccessibilityRules_SecondClassBody()
        {
            return
@"
Private memberString As String
Private memberLong As Long
Public mySecondEggo As String


Public Sub Foo2( arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
    memberString = arg1 & ""Foo""
End Sub

Public Function FooBar(ByRef arg1 As String, theSecondArg As Long) As Long
    Let arg1 = ""test""
    Dim result As String
    result = Clg(arg1) + theSecondArg
    FooBar = result
End Function

Property Let GoMyOtherEggo(newValue As String)
    mySecondEggo = newValue
End Property

Property Get GoMyOtherEggo()
    GoMyOtherEggo = mySecondEggo
End Property

Private Sub FooFighters(ByRef arg1 As String)
    xArg1 = 6
    Let arg1 = ""test""
End Sub

Sub Bar()
    Dim st As String
    st = ""Test""
    Dim v As Long
    v = 5
    result = FooBar(st, v)
End Sub
";
        }

        private string GetRespectsDeclarationsAccessibilityRules_FirstModuleBody()
        {
            return
@"
Option Explicit


Public Const CantTouchThis As String = ""Can't Touch this""
Public THE_FILENAME As String

Sub SetFilename(filename As String)
    implicitVar = 7
    THE_FILENAME = filename
End Sub
";
        }

        private string GetRespectsDeclarationsAccessibilityRules_SecondModuleBody()
        {
            return
@"
Option Explicit


Public BigNumber as Long
Public ShortStory As String

Public Sub DoSomething(filename As String)
    ShortStory = filename
End Sub
";
        }
        #endregion

        [TestCategory("Resolver")]
        [TestMethod]
        public void DeclarationFinderCanCopeWithMultipleModulesImplementingTheSameInterface()
        {
            var project = GetTestProject("testProject");
            var interf = GetTestClassModule(project, "interface");
            var member = GetTestFunction(interf, "testMember", Accessibility.Public);
            var implementingClass1 = GetTestClassModule(project, "implementingClass1");
            var implementingClass2 = GetTestClassModule(project, "implementingClass2");
            var implementsContext1 = new VBAParser.ImplementsStmtContext(null, 0);
            var implementsContext2 = new VBAParser.ImplementsStmtContext(null, 0);
            AddReference(interf, implementingClass1, implementsContext1);
            AddReference(interf, implementingClass1, implementsContext2);
            var declarations = new List<Declaration> {interf, member, implementingClass1, implementingClass2};

            DeclarationFinder finder = new DeclarationFinder(declarations, new List<Rubberduck.Parsing.Annotations.IAnnotation>(), new List<UnboundMemberDeclaration>());
            var interfaceDeclarations = finder.FindAllInterfaceMembers().ToList();

            Assert.AreEqual(1, interfaceDeclarations.Count());
        }

        private static ClassModuleDeclaration GetTestClassModule(Declaration projectDeclatation, string name, bool isExposed = false)
        {
            var qualifiedClassModuleMemberName = new QualifiedMemberName(StubQualifiedModuleName(name), name);
            var classModuleAttributes = new Rubberduck.Parsing.VBA.Attributes();
            if (isExposed)
            {
                classModuleAttributes.AddExposedClassAttribute();
            }
            return new ClassModuleDeclaration(qualifiedClassModuleMemberName, projectDeclatation, name, false, null, classModuleAttributes);
        }

        private static ProjectDeclaration GetTestProject(string name)
        {
            var qualifiedProjectName = new QualifiedMemberName(StubQualifiedModuleName("proj"), name);
            return new ProjectDeclaration(qualifiedProjectName, name, false, null);
        }

        private static QualifiedModuleName StubQualifiedModuleName(string name)
        {
            return new QualifiedModuleName("dummy", "dummy", name);
        }

        private static FunctionDeclaration GetTestFunction(Declaration moduleDeclatation, string name, Accessibility functionAccessibility)
        {
            var qualifiedFunctionMemberName = new QualifiedMemberName(moduleDeclatation.QualifiedName.QualifiedModuleName, name);
            return new FunctionDeclaration(qualifiedFunctionMemberName, moduleDeclatation, moduleDeclatation, "test", null, "test", functionAccessibility, null, Selection.Home, false, false, null, null);
        }

        private static void AddReference(Declaration toDeclaration, Declaration fromModuleDeclaration, ParserRuleContext context = null)
        {
            toDeclaration.AddReference(toDeclaration.QualifiedName.QualifiedModuleName, fromModuleDeclaration, fromModuleDeclaration, context, toDeclaration.IdentifierName, toDeclaration, Selection.Home, new List<Rubberduck.Parsing.Annotations.IAnnotation>());
        }
    }
}

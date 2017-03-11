
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using Rubberduck.VBEditor;
using Antlr4.Runtime;
using Rubberduck.Inspections;

namespace RubberduckTests.Symbols
{
    [TestClass]
    public class DeclarationFinderTests
    {
        private AccessibilityTestsDataObject _tdo;

        [TestMethod]
        [TestCategory("Resolver")]
        public void DeclarationFinder_FindsAccessibleDeclarations_InProcedure()
        {
            SetupSUT("TestProject");
            var scopeName = "ProcedureScope";
            var names = new List<string>();
            _tdo.AccessibleNames.TryGetValue(scopeName, out names);
            TestAccessibleDeclarations(names, scopeName, true, false);
        }

        [TestMethod]
        [TestCategory("Resolver")]
        public void DeclarationFinder_FindsAccessibleDeclarations_ModuleScope()
        {
            SetupSUT("TestProject");
            var scopeName = "ModuleScope";
            var names = new List<string>();
            _tdo.AccessibleNames.TryGetValue(scopeName, out names);
            TestAccessibleDeclarations(names, scopeName, true, false);
        }

        [TestMethod]
        [TestCategory("Resolver")]
        public void DeclarationFinder_FindsAccessibleDeclarations_GlobalScope()
        {
            SetupSUT("TestProject");
            var scopeName = "GlobalScope";
            var names = new List<string>();
            _tdo.AccessibleNames.TryGetValue(scopeName, out names);
            TestAccessibleDeclarations(names, scopeName, true, true);
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
            var allNamesUsedInTests = new List<string>();
            List<string> names;
            foreach( var Key in _tdo.AccessibleNames.Keys)
            {
                _tdo.AccessibleNames.TryGetValue(Key, out names);
                allNamesUsedInTests.AddRange(names);
            }
            allNamesUsedInTests.AddRange(_tdo.Components.Select(n => n.Name));

            var target = GetTargetForAccessibilityTests();

            var accessibleNames =
                _tdo.Parser.State.DeclarationFinder.GetDeclarationsWithIdentifiersToAvoid(target)
                    .Select( dec => dec.IdentifierName);

            Assert.IsTrue(accessibleNames.Count() > 0);

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
                allIdentifiersToCheck.AddRange(_tdo.Components.Select(n => n.Name));
            }

            var target = GetTargetForAccessibilityTests();

            var declarationFinderResults =
                _tdo.Parser.State.DeclarationFinder.GetDeclarationsWithIdentifiersToAvoid(target);

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
            var targets = _tdo.Parser.State.AllUserDeclarations.Where(dec => dec.IdentifierName == _tdo.TargetIdentifier);
            if (targets.Count() > 1)
            {
                Assert.Inconclusive("Multiple targets found with identifier: " + _tdo.TargetIdentifier + ".  Test requires a unique identifierName");
            }
            return targets.FirstOrDefault();
        }

        private void SetupSUT(string projectName)
        {
            if (_tdo != null) { return; }

            string[] accessibleWithinParentProcedure = { "arg1", "FooBar1", "targetAccessibilityTests", "theSecondArg" };
            string[] accessibleModuleScope = { "memberString", "memberLong", "myEggo", "Foo", "FooBar1", "GoMyEggo", "FooFight" };
            string[] accessibleGlobalScope = { "CantTouchThis", "BigNumber", "DoSomething", "SetFilename", "ShortStory","THE_FILENAME", "TestProject"};

            var firstClassBody = FindsAccessibleDeclarations_FirstClassBody();
            var secondClassBody = FindsAccessibleDeclarations_SecondClassBody();
            var firstModuleBody = FindsAccessibleDeclarations_FirstModuleBody();
            var secondModuleBody = FindsAccessibleDeclarations_SecondModuleBody();

            _tdo = new AccessibilityTestsDataObject();
            _tdo.TargetIdentifier = "targetAccessibilityTests";
            AddAccessibleNames("ProcedureScope", accessibleWithinParentProcedure);
            AddAccessibleNames("ModuleScope", accessibleModuleScope);
            AddAccessibleNames("GlobalScope", accessibleGlobalScope);

            AddTestComponent("CFirstClass", firstClassBody, ComponentType.ClassModule);
            AddTestComponent("CSecondClass", secondClassBody, ComponentType.ClassModule);
            AddTestComponent("modFirst", firstModuleBody, ComponentType.StandardModule);
            AddTestComponent("modSecond", secondModuleBody, ComponentType.StandardModule);

            _tdo.VBE = BuildProject("TestProject", _tdo.Components);

            _tdo.Parser = MockParser.Create(_tdo.VBE, new RubberduckParserState(_tdo.VBE));
            _tdo.Parser.Parse(new CancellationTokenSource());
            if (_tdo.Parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }
        }

        private void AddAccessibleNames( string scope, string[] accessibleNames)
        {
            if(null == _tdo.AccessibleNames)
            {
                _tdo.AccessibleNames = new Dictionary<string, List<string>>();
            }
            _tdo.AccessibleNames.Add(scope, accessibleNames.ToList());
        }

        private void AddTestComponent( string moduleIdentifier, string moduleContent, ComponentType componentType)
        {
            if( null ==_tdo.Components)
            {
                _tdo.Components = new List<TestComponentSpecification>();
            }
            _tdo.Components.Add(new TestComponentSpecification(moduleIdentifier, moduleContent, componentType));
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

        internal class AccessibilityTestsDataObject
        {
            public IVBE VBE { get; set; }
            public ParseCoordinator Parser { get; set; }
            public List<TestComponentSpecification> Components { get; set; }
            public Dictionary<string,List<string>> AccessibleNames { get; set;  }
            public string TargetIdentifier { get; set; }
        }

        #region AccessibilityTestsModuleContent
        private string FindsAccessibleDeclarations_FirstClassBody()
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

        private string FindsAccessibleDeclarations_SecondClassBody()
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

        private string FindsAccessibleDeclarations_FirstModuleBody()
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

        private string FindsAccessibleDeclarations_SecondModuleBody()
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

        [TestMethod]
        [Ignore] // ref. https://github.com/rubberduck-vba/Rubberduck/issues/2330
        public void FiendishlyAmbiguousNameSelectsSmallestScopedDeclaration()
        {
            var code = @"
Option Explicit

Public Sub foo()
    Dim foo As Long
    foo = 42
    Debug.Print foo
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("foo", ProjectProtection.Unprotected)
                .AddComponent("foo", ComponentType.StandardModule, code, new Selection(6, 6))
                .MockVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.Parse(new CancellationTokenSource());

            var expected = parser.State.AllDeclarations.Single(item => item.DeclarationType == DeclarationType.Variable);
            var actual = parser.State.DeclarationFinder.FindSelectedDeclaration(vbe.Object.ActiveCodePane);

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected.DeclarationType, actual.DeclarationType);
        }

        [TestMethod]
        [Ignore] // bug: this test should pass... it's not all that evil
        public void AmbiguousNameSelectsSmallestScopedDeclaration()
        {
            var code = @"
Option Explicit

Public Sub foo()
    Dim foo As Long
    foo = 42
    Debug.Print foo
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code, new Selection(6, 6))
                .MockVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.Parse(new CancellationTokenSource());

            var expected = parser.State.AllDeclarations.Single(item => item.DeclarationType == DeclarationType.Variable);
            var actual = parser.State.DeclarationFinder.FindSelectedDeclaration(vbe.Object.ActiveCodePane);

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected.DeclarationType, actual.DeclarationType);
        }

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
            var declarations = new List<Declaration> { interf, member, implementingClass1, implementingClass2 };

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
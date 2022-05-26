using NUnit.Framework;
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
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using System;

namespace RubberduckTests.Symbols
{

    [TestFixture]
    public class DeclarationFinderTests
    {
        [TestCase("member1", true)]
        [TestCase("adder", true)]
        [TestCase("Foo", false)]
        [Category("Resolver")]
        public void DeclarationFinder_InProcedure_MethodDeclaration(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"

Private member1 As Long

Public Function F|oo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Foo = member1
End Function
";

            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = "CFirstClass"
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.ClassModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck, conflicts));
        }

        [TestCase("member1", true)]
        [TestCase("adder", false)]
        [TestCase("Foo", true)]
        [Category("Resolver")]
        public void DeclarationFinder_InProcedure_LocalVariableReferences(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"

Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + ad|der
    Foo = member1
End Function
";

            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = "CFirstClass"
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.ClassModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck, conflicts));
        }

        [TestCase("member1", false)]
        [TestCase("adder", true)]
        [TestCase("Foo", true)]
        [Category("Resolver")]
        public void DeclarationFinder_InProcedure_MemberDeclaration(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
@"
Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Foo = m|ember1
End Function
";
            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = "CFirstClass"
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.ClassModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck, conflicts));
        }

        [TestCase("member1", false)]
        [TestCase("member2", false)]
        [TestCase("adder", true)]
        [TestCase("Foo", true)]
        [TestCase("Foo2", false)]
        [TestCase("Bar", false)]
        [Category("Resolver")]
        public void DeclarationFinder_ModuleScope(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"

Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = memb|er1 + adder
    Foo = member1
End Function
";
            var moduleContent2 =
            @"

Private member1 As Long
Public member2 As Long

Public Function Foo2() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Foo2 = member1
End Function

Private Sub Bar()
    member2 = member2 * 4
End Sub
";

            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = "CFirstClass"
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.ClassModule);
            AddTestComponent(tdo, "modOne", moduleContent2, ComponentType.StandardModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck,conflicts)); 
        }

        [TestCase("Foo", false)]
        [TestCase("Foo2", false)]
        [Category("Resolver")]
        public void DeclarationFinder_PublicClassAndPublicModuleSub_RenameClassSub(string nameToCheck, bool isConflict)
        {
            var moduleContent1 = 
            @"
Public Function Fo|o() As Long
    Foo = 5
End Function
";
            var moduleContent2 =
            @"
Public Function Foo2() As Long
    Foo2 = 2
End Function
";

            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = "CFirstClass"
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.ClassModule);
            AddTestComponent(tdo, "modOne", moduleContent2, ComponentType.StandardModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck, conflicts));
        }

        [Test]
        [TestCase("Foo", false)]
        [TestCase("Foo2", true)]
        [TestCase("member11", true)]
        [TestCase("member1", false)]
        [TestCase("Bar", true)]
        [TestCase("adder", false)]
        [Category("Resolver")]
        public void DeclarationFinder_Module_To_ClassScope(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
            @"

Private member11 As Long
Public member2 As Long

Public Function Foo2() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Foo2 = member1
End Function

Private Sub Bar()
    member2 = membe|r2 * 4
End Sub
";
            var moduleContent2 =
            @"

Private member1 As Long

Public Function Foo() As Long
    Dim adder as Long
    adder = 10
    member1 = member1 + adder
    Foo = member1
End Function
";

            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = "modOne"
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.StandardModule);
            AddTestComponent(tdo, "CFirstClass", moduleContent2, ComponentType.ClassModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck, conflicts));
        }

        [TestCase("DoThis", true)]
        [TestCase("member1", true)]
        [TestCase("filename", true)]
        [Category("Resolver")]
        public void DeclarationFinder_PrivateSub_CheckConflictsInOtherModules(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
@"
Private Sub DoThis(filename As String)
    SetFi|lename filename
End Sub
";
            var moduleContent2 =
@"
Private member1 As String

Public Sub SetFilename(filename As String)
    member1 = filename
End Sub
";

            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = "modOne"
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.StandardModule);
            AddTestComponent(tdo, "modTwo", moduleContent2, ComponentType.StandardModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck, conflicts));
        }

        [TestCase("DoThis", true)]
        [TestCase("filename", true)]
        [TestCase("member1", true)]
        [TestCase("mFolderpath", true)]
        [TestCase("ExtractFilename", true)]
        [TestCase("StoreFilename", true)]
        [TestCase("filepath", true)]
        [Category("Resolver")]
        public void DeclarationFinder_PrivateSub_MultipleReferences(string nameToCheck, bool isConflict)
        {

            var moduleContent1 =
@"
Private Sub DoThis(filename As String)
    SetFil|ename filename
End Sub
";
            var moduleContent2 =
@"
Private member1 As String

Public Sub SetFilename(filename As String)
    member1 = filename
End Sub
";
            var moduleContent3 =
@"
Private mFolderpath As String

Private Sub StoreFilename(filepath As String)
    Dim filename As String
    filename = ExtractFilename(filepath)
    SetFilename filename
End Sub

Private Function ExtractFilename(filepath As String) As String
    ExtractFilename = filepath
End Function"
;

            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = "modOne"
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.StandardModule);
            AddTestComponent(tdo, "modTwo", moduleContent2, ComponentType.StandardModule);
            AddTestComponent(tdo, "modThree", moduleContent3, ComponentType.StandardModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck, conflicts));
        }

        //https://github.com/rubberduck-vba/Rubberduck/issues/5694
        [Test]
        [Category("Refactorings")]
        [Category("Resolver")]
        public void RenameRefactoring_RenameClassQualifiedMember_NoConflict()
        {
            var isConflict = false;
            var nameToCheck = "ThisIsMyProperty";
            var moduleContent1 =
@"
Option Explicit

Private refdClass As ReferencedClass

Private Sub Class_Initialize()
    Set refdClass = New ReferencedClass
End Sub

Public Function ThisIsM|yFunc() As Long
    ThisIsMyFunc = refdClass.ThisIsMyProperty
End Function
";

            var moduleContent2 =
@"
Option Explicit

Public Property Get ThisIsMyProperty() As Long
    ThisIsMyProperty = 5
End Property        
";

            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = MockVbeBuilder.TestModuleName
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.ClassModule);
            AddTestComponent(tdo, "ReferencedClass", moduleContent2, ComponentType.ClassModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck, conflicts));
        }

        [Test]
        [Category("Refactorings")]
        [Category("Resolver")]
        public void RenameRefactoring_RenameToUnQualifiedModuleMember_HasConflict()
        {
            var isConflict = true;
            var nameToCheck = "ThisIsMyProperty";
            var moduleContent1 =
@"
Option Explicit

Public Function ThisIsM|yFunc() As Long
    ThisIsMyFunc = ThisIsMyProperty
End Function
";

            var moduleContent2 =
@"
Option Explicit

Public Property Get ThisIsMyProperty() As Long
    ThisIsMyProperty = 5
End Property        
";

            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = MockVbeBuilder.TestModuleName
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.ClassModule);
            AddTestComponent(tdo, "ReferencedClass", moduleContent2, ComponentType.StandardModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck, conflicts));
        }

        [TestCase("Bar", true)]
        [TestCase("myData", true)]
        [TestCase("mDupData", true)]
        [TestCase("filepath", true)]
        [TestCase("helper", true)]
        [TestCase("CFileHelper", true)]
        [TestCase("filename", true)]
        [TestCase("mFolderpath", true)]
        [TestCase("ExtractFilename", true)]
        [TestCase("SetFilename", true)]
        [TestCase("Foo", false)]
        [TestCase("FooBar", false)]
        [Category("Resolver")]
        public void DeclarationFinder_PrivateSub_WithBlock(string nameToCheck, bool isConflict)
        {
            var moduleContent1 =
@"
Private myData As String
Private mDupData As String

Public Sub Foo(filenm As String)
    Dim filepath As String
    filepath = ""C:\MyStuff\"" & filenm
    Dim helper As CFileHelper
    Set helper = new CFileHelper
    With helper
        .StoreFile|name filepath
        mDupData = filepath
    End With
End Sub

Public Sub StoreFilename(filename As String)
    myData = filename
End Sub

Private Sub FooBar()
End Sub
";
            var moduleContent2 =
@"
Private mFolderpath As String

Public Sub StoreFilename(input As String)
    Dim filename As String
    filename = ExtractFilename(input)
    SetFilename filename
End Sub

Private Function ExtractFilename(filepath As String) As String
    ExtractFilename = filepath
End Function

Public Sub Bar()
End Sub
";

            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = "modOne"
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.StandardModule);
            AddTestComponent(tdo, "CFileHelper", moduleContent2, ComponentType.ClassModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck, conflicts));
        }

        [TestCase("Foo1", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Foo2", false, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Foo3", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Foo4", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("gConstant", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("member2", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("member11", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("gConstant", true, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Bar1", false, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Bar1", true, "Foo1 + Fo|o2 + Foo3")]
        [TestCase("Bar2", false, "modTwo.Foo1 + modTwo.Fo|o2 + modTwo.Foo3")]
        [TestCase("Bar2", true, "Foo1 + Fo|o2 + Foo3")]
        [Category("Resolver")]
        public void DeclarationFinder_Module_To_ModuleScopeResolution(string nameToCheck, bool isConflict, string scopeResolvedInput)
        {
            var moduleContent1 =
$@"
Private member11 As Long
Public member2 As Long

Private Function Bar1() As Long
    Bar2
    Bar1 = member2 + {scopeResolvedInput}
End Function

Private Sub Bar2()
    member2 = member2 * 4 
End Sub
";
            var moduleContent2 =
@"
Public Const gConstant As Long = 10

Public Function Foo1() As Long
    Foo1 = 1
End Function

Public Function Foo2() As Long
    Foo2 = 2
End Function

Public Function Foo3() As Long
    Foo3 = 3
End Function

Private Sub Foo4()
End Sub
";

            var tdo = new AccessibilityTestsDataObject(moduleContent1)
            {
                SelectionModuleName = "modOne"
            };

            AddTestComponent(tdo, tdo.SelectionModuleName, ComponentType.StandardModule);
            AddTestComponent(tdo, "modTwo", moduleContent2, ComponentType.StandardModule);
            var conflicts = TestConflictingDeclaration(tdo, nameToCheck);
            Assert.AreEqual(isConflict, conflicts.Where(cf => cf.IdentifierName.Equals(nameToCheck)).Any(), ConflictMessage(isConflict, nameToCheck, conflicts));
        }


        //https://github.com/rubberduck-vba/Rubberduck/issues/4969
        private const string projectOneModuleName = "projectOneModule";
        private const string projectTwoModuleName = "projectTwoModule";
        [TestCase(projectOneModuleName, 0)]  //Duplicate module name found in a separate project
        [TestCase(projectTwoModuleName, 1)] //Duplicate module name found in the same project
        [Category("Resolver")]
        public void DeclarationFinder_NameConflictDetectionRespectsProjectScope(string proposedTestModuleName, int expectedCount)
        {

            string renameTargetModuleName = "TargetModule";

            string moduleContent = $"Private Sub Foo(){Environment.NewLine}End Sub";

            var projectOneContent = new TestComponentSpecification[]
            {
                new TestComponentSpecification(projectOneModuleName, moduleContent, ComponentType.StandardModule)
            };

            var projectTwoContent = new TestComponentSpecification[]
            {
                new TestComponentSpecification(renameTargetModuleName, moduleContent, ComponentType.StandardModule),
                new TestComponentSpecification(projectTwoModuleName, moduleContent, ComponentType.StandardModule)
            };

            var vbe = BuildProjects(new (string, IEnumerable<TestComponentSpecification>)[]
                {("ProjectOne", projectOneContent),("ProjectTwo", projectTwoContent)});

            using(var parser = MockParser.CreateAndParse(vbe))
            {
                var target = parser.DeclarationFinder.UserDeclarations(DeclarationType.ProceduralModule)
                    .FirstOrDefault(item => item.IdentifierName.Equals(renameTargetModuleName));

                var results = parser.DeclarationFinder.FindNewDeclarationNameConflicts(proposedTestModuleName, target);

                Assert.AreEqual(expectedCount, results.Count());
            }
        }

        private static string ConflictMessage(bool isConflict, string name, IEnumerable<Declaration> conflicts)
        {
            return isConflict ? $"Identifier '{name}' is a conflict but was not identified" : $"Identifier '{name}' was incorrectly found as a conflict";
        }

        private IEnumerable<Declaration> TestConflictingDeclaration(AccessibilityTestsDataObject tdo, string name)
        {

            tdo.VBE = BuildProject(tdo.ProjectName, tdo.Components);
            tdo.Parser = MockParser.CreateAndParse(tdo.VBE);
            PrepareScenarioTestData(tdo, name);

            AcquireTarget(tdo, out Declaration target, tdo.QualifiedSelection);
            return tdo.Parser.DeclarationFinder.FindNewDeclarationNameConflicts(name, target);
        }

        private void AcquireTarget(AccessibilityTestsDataObject tdo, out Declaration target, QualifiedSelection selection)
        {
            target = tdo.Parser.DeclarationFinder.AllDeclarations
                .Where(item => item.IsUserDefined)
                .FirstOrDefault(item => item.IsSelected(selection) || item.References.Any(r => r.IsSelected(selection)));
        }


        private void PrepareScenarioTestData(AccessibilityTestsDataObject tdo, string name)
        {
            tdo.VBE = BuildProject(tdo.ProjectName, tdo.Components);
            tdo.Parser = MockParser.CreateAndParse(tdo.VBE);

            var component = RetrieveComponent(tdo, tdo.SelectionModuleName);
            var moduleParent = component.CodeModule.Parent;
            tdo.QualifiedSelection = new QualifiedSelection(new QualifiedModuleName(moduleParent), tdo.Target);
        }

        private void AddTestComponent(AccessibilityTestsDataObject tdo, string moduleIdentifier, ComponentType componentType)
        {
            if (null == tdo.Components)
            {
                tdo.Components = new List<TestComponentSpecification>();
            }
            tdo.Components.Add(new TestComponentSpecification(moduleIdentifier, tdo.Code, componentType));
        }

        private void AddTestComponent(AccessibilityTestsDataObject tdo, string moduleIdentifier, string moduleContent, ComponentType componentType)
        {
            if (null == tdo.Components)
            {
                tdo.Components = new List<TestComponentSpecification>();
            }
            tdo.Components.Add(new TestComponentSpecification(moduleIdentifier, moduleContent, componentType));
        }

        private IVBE BuildProject(string projectName, List<TestComponentSpecification> testComponents)
        {
            var projectDefs = new (string, IEnumerable<TestComponentSpecification>)[] { (projectName, testComponents) };
            return BuildProjects(projectDefs);
        }

        private IVBE BuildProjects(IEnumerable<(string ProjectName, IEnumerable<TestComponentSpecification> TestComponents)> projectDefinitions)
        {
            var builder = new MockVbeBuilder();
            foreach (var projectDef in projectDefinitions)
            {
                builder = AddProject(builder, projectDef.ProjectName, projectDef.TestComponents);
            }
            return builder.Build().Object;
        }

        private MockVbeBuilder AddProject(MockVbeBuilder builder, string projectName, IEnumerable<TestComponentSpecification> testComponents)
        {
            var enclosingProjectBuilder = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected);

            foreach (var testComponent in testComponents)
            {
                enclosingProjectBuilder.AddComponent(testComponent.Name, testComponent.ModuleType, testComponent.Content);
            }

            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            return builder;
        }

        private IVBComponent RetrieveComponent(AccessibilityTestsDataObject tdo, string componentName)
        {
            var vbProject = tdo.VBE.VBProjects.Where(item => item.Name == tdo.ProjectName).SingleOrDefault();
            return vbProject.VBComponents.Where(item => item.Name == componentName).SingleOrDefault();
        }

        internal class TestComponentSpecification
        {
            public TestComponentSpecification(string componentName, string componentContent, ComponentType componentType)
            {
                Name = componentName;
                Content = componentContent;
                ModuleType = componentType;
            }

            public string Name { set;  get; }
            public string Content { set; get; }
            public ComponentType ModuleType { set; get; }
        }


        internal class AccessibilityTestsDataObject
        {
            private CodeString _codeString;
            public AccessibilityTestsDataObject()
            {
                ProjectName = "TestProject";
            }
            public AccessibilityTestsDataObject(string moduleCode)
            {
                ProjectName = "TestProject";
                _codeString = moduleCode.ToCodeString();
            }
            public IVBE VBE { get; set; }
            public RubberduckParserState Parser { get; set; }
            public List<TestComponentSpecification> Components { get; set; }
            public string ProjectName { get; set; }
            public string SelectionModuleName { get; set; }
            public string SelectionTarget { get; set; }
            public string SelectionLineIdentifier { get; set; }
            public List<string> ExpectedResults { get; set; }
            public QualifiedSelection QualifiedSelection { get; set; }
            public string Code => _codeString.Code;
            public Selection Target => _codeString.CaretPosition.ToOneBased();
        }

        [Test]
        [Category("Resolver")]
        public void SameNameForProjectAndClassImplicit_ScopedDeclaration()
        {
            var refEditClass = @"
Option Explicit

Private ValueField As Variant

Public Property Get Value()
  Value = ValueField
End Property

Public Property Let Value(Value As Variant)
  ValueField = Value
End Property";

            var code =
                @"
Option Explicit

Public Sub foo()
    Dim myEdit As RefEdit
    Set myEdit = New RefEdit

    myEdit.Value = ""abc""
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("RefEdit", ProjectProtection.Unprotected)
                .AddComponent("RefEdit", ComponentType.ClassModule, refEditClass)
                .AddComponent("Test", ComponentType.StandardModule, code, new Selection(7, 6))
                .AddProjectToVbeBuilder()
                .Build();

            var parser = MockParser.Create(vbe.Object);
            parser.Parse(new CancellationTokenSource());

            var expected = ParserState.ResolverError;
            var actual = parser.State.Status;

            Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
        }

        [Category("Resolver")]
        [Test]
        public void Identify_NamedParameter_Parameter_FromExcel_DefaultAccess()
        {
            // Note that ColumnIndex is actually a parameter of the _Default default member
            // of the Excel.Range object.
            const string code = @"
Public Sub DoIt()
    Dim foo As Variant
    Dim sht As WorkSheet

    foo = sht.Cells(ColumnIndex:=12).Value
End Sub";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("TestModule", ComponentType.StandardModule, code)
                .AddReference(ReferenceLibrary.Excel)
                .AddProjectToVbeBuilder()
                .Build();

            var selection = new Selection(6, 21, 6, 32);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var module = state.DeclarationFinder.AllModules.First(qmn => qmn.ComponentName.Equals("TestModule"));
                var qualifiedSelection = new QualifiedSelection(module, selection);

                var reference = state.DeclarationFinder.IdentifierReferences(qualifiedSelection).First();
                var referencedDeclaration = reference.Declaration;

                var expectedReferencedDeclarationName = "EXCEL.EXE;Excel.Range._Default.Let.ColumnIndex";
                var actualReferencedDeclarationName = $"{referencedDeclaration.ParentScope}.{referencedDeclaration.IdentifierName}";

                Assert.AreEqual(expectedReferencedDeclarationName, actualReferencedDeclarationName);
                Assert.AreEqual(DeclarationType.Parameter, referencedDeclaration.DeclarationType);
            }
        }

        [Test]
        [Category("Resolver")]
        public void FindParameterFromArgument_WorksWithMultipleScopes()
        {
            var module1 =
@"Public Sub Foo(arg As Variant)
End Sub";

            var module2 =
@"Private Sub Foo(expected As Variant)
End Sub

Public Sub Bar()
    Dim fooBar As Variant
    Foo fooBar
End Sub
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, module1, new Selection(1, 1))
                .AddComponent("Module2", ComponentType.StandardModule, module2, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.FirstOrDefault(decl => decl.IdentifierName.Equals("expected"));

                var enclosing = declarations.FirstOrDefault(decl => decl.IdentifierName.Equals("Bar"));
                var context = enclosing?.Context.GetDescendent<VBAParser.ArgumentContext>();
                var actual = state.DeclarationFinder.FindParameterOfNonDefaultMemberFromSimpleArgumentNotPassedByValExplicitly(context, enclosing);

                Assert.AreEqual(expected, actual);
            }
        }


        [Category("Resolver")]
        [Category("Interfaces")]
        [Test]
        public void DeclarationFinderCanCopeWithMultipleModulesImplementingTheSameInterface()
        {
            const string interfaceCode = @"
Public Sub Foo()
End Sub
";

            const string implementationCode = @"
Implements IClass1

Public Sub IClass1_Foo()
End Sub
";

            var vbe = new MockVbeBuilder()
                .ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("IClass1", ComponentType.ClassModule, interfaceCode, new Selection(0, 0))
                .AddComponent("Class2", ComponentType.ClassModule, implementationCode, new Selection(0, 0))
                .AddComponent("Class3", ComponentType.ClassModule, implementationCode, new Selection(0, 0))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var interfaceDeclarations = state.DeclarationFinder.FindAllInterfaceMembers().ToList();

                Assert.AreEqual(1, interfaceDeclarations.Count());
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindInterfaceImplementationMembersForGetMatchesOnlyGet()
        {
            var intrface =
@"Option Explicit

Public Property Get Foo(Bar As Long) As Long
End Property

Public Property Let Foo(Bar As Long, NewValue As Long)
End Property
";

            var implementation =
@"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo(Bar As Long) As Long
End Property

Public Property Let TestInterface_Foo(Bar As Long, RHS As Long)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("Foo"));

                var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("TestInterface_Foo"));
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(1, results, "Expected {0} Declarations, received {1}", expected, results);
                Assert.AreEqual(expected, actual.First(), "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindInterfaceImplementationMembersForLetMatchesOnlyLet()
        {
            var intrface =
                @"Option Explicit

Public Property Get Foo(Bar As Long) As Long
End Property

Public Property Let Foo(Bar As Long, NewValue As Long)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo(Bar As Long) As Long
End Property

Public Property Let TestInterface_Foo(Bar As Long, RHS As Long)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyLet && decl.IdentifierName.Equals("Foo"));

                var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyLet && decl.IdentifierName.Equals("TestInterface_Foo"));
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(1, results, "Expected {0} Declarations, received {1}", expected, results);
                Assert.AreEqual(expected, actual.First(), "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindInterfaceImplementationMembersForSetMatchesOnlySet()
        {
            var intrface =
                @"Option Explicit

Public Property Get Foo(Bar As Long) As Object
End Property

Public Property Set Foo(Bar As Long, NewValue As Object)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo(Bar As Long) As Object
End Property

Public Property Set TestInterface_Foo(Bar As Long, RHS As Object)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertySet && decl.IdentifierName.Equals("Foo"));

                var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertySet && decl.IdentifierName.Equals("TestInterface_Foo"));
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(1, results, "Expected {0} Declarations, received {1}", expected, results);
                Assert.AreEqual(expected, actual.First(), "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindInterfaceMembersMatchesPublicVariables()
        {
            var intrface =
                @"Option Explicit

Public Bar As String

Public Property Get Foo() As Long
End Property

Public Property Let Foo(rhs As Long)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Bar() As String
End Property

Public Property Let TestInterface_Bar(rhs As String)
End Property

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Let TestInterface_Foo(rhs As Long)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var actual = state.DeclarationFinder.FindAllInterfaceMembers().ToList();
                var expected = state.DeclarationFinder.AllUserDeclarations.Where(decl => decl.ParentScope.Equals("UnderTest.TestInterface")).ToList();

                var results = actual.Count;

                Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
                Assert.That(actual, Is.EquivalentTo(expected));
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindInterfaceImplementationMembersMatchesPublicVariables()
        {
            var intrface =
                @"Option Explicit

Public Bar As String

Public Property Get Foo() As Long
End Property

Public Property Let Foo(rhs As Long)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Bar() As String
End Property

Public Property Let TestInterface_Bar(rhs As String)
End Property

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Let TestInterface_Foo(rhs As Long)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

                var expected = declarations.Where(decl => decl.DeclarationType.HasFlag(DeclarationType.Property) && decl.IdentifierName.Equals("TestInterface_Bar")).ToList();
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
                Assert.That(actual, Is.EquivalentTo(expected));
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindInterfaceImplementationMembersPublicVariantMatchesAllPropertyTypes()
        {
            var intrface =
                @"Option Explicit

Public Bar As Variant";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Bar() As Variant
End Property

Public Property Let TestInterface_Bar(rhs As Variant)
End Property

Public Property Set TestInterface_Bar(rhs As Variant)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

                var expected = declarations.Where(decl => decl.DeclarationType.HasFlag(DeclarationType.Property) && decl.IdentifierName.Equals("TestInterface_Bar")).ToList();
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
                Assert.That(actual, Is.EquivalentTo(expected));
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindInterfaceImplementationMembersPublicIntrinsicDoesNotMatchSet()
        {
            var intrface =
                @"Option Explicit

Public Bar As Long";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Bar() As Long
End Property

Public Property Let TestInterface_Bar(rhs As Long)
End Property

Public Property Set TestInterface_Bar(rhs As Variant)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

                var expected = declarations.Where(decl => decl.IdentifierName.Equals("TestInterface_Bar") &&
                                                          decl.DeclarationType == DeclarationType.PropertyLet ||
                                                          decl.DeclarationType == DeclarationType.PropertyGet).ToList();
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
                Assert.That(actual, Is.EquivalentTo(expected));
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindInterfaceImplementationMembersPublicObjectDoesNotMatchLet()
        {
            var intrface =
                @"Option Explicit

Public Bar As Object";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Bar() As Object
End Property

Public Property Let TestInterface_Bar(rhs As Variant)
End Property

Public Property Set TestInterface_Bar(rhs As Object)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.Variable && decl.IdentifierName.Equals("Bar"));

                var expected = declarations.Where(decl => decl.IdentifierName.Equals("TestInterface_Bar") &&
                                                          decl.DeclarationType == DeclarationType.PropertySet ||
                                                          decl.DeclarationType == DeclarationType.PropertyGet).ToList();
                var actual = state.DeclarationFinder.FindInterfaceImplementationMembers(declaration).ToList();
                var results = actual.Count;

                Assert.AreEqual(expected.Count, results, "Expected {0} Declarations, received {1}", expected.Count, results);
                Assert.That(actual, Is.EquivalentTo(expected));
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindFindInterfaceMemberMatchesDeclarationTypes()
        {
            var intrface =
                @"Option Explicit

Public Property Get Foo(Bar As Long) As Long
End Property

Public Property Let Foo(Bar As Long, NewValue As Long)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo(Bar As Long) As Long
End Property

Public Property Let TestInterface_Foo(Bar As Long, RHS As Long)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("TestInterface_Foo"));

                var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("Foo"));
                var actual = (declaration as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindFindInterfaceMemberParameterNamesIgnored()
        {
            var intrface =
                @"Option Explicit

Public Property Get Foo(Bar As Long) As Long
End Property

Public Property Let Foo(Bar As Long, Baz As Long)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo(Bar As Long) As Long
End Property

Public Property Let TestInterface_Foo(Bar As Long, RHS As Long)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyLet && decl.IdentifierName.Equals("TestInterface_Foo"));

                var expected = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyLet && decl.IdentifierName.Equals("Foo"));
                var actual = (declaration as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void FindFindInterfaceMemberNoResultWithoutMatchingDeclaration()
        {
            var intrface =
                @"Option Explicit

Public Property Let Foo(Bar As Long, NewValue As Long)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo(Bar As Long) As Long
End Property

Public Property Let TestInterface_Foo(Bar As Long, RHS As Long)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var declaration = declarations.Single(decl => decl.DeclarationType == DeclarationType.PropertyGet && decl.IdentifierName.Equals("TestInterface_Foo"));

                var actual = (declaration as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.IsNull(actual, "Expected null, resolved to {0}", actual);
            }
        }

        [Test]
        [Category("Resolver")]
        public void ImplementsInterfaceMemberMatchesProperty()
        {
            var intrface =
                @"Option Explicit

Public Property Get Foo() As Long
End Property

Public Property Let Foo(rhs As Long)
End Property
";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Let TestInterface_Foo(rhs As Long)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.PropertyGet);
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyGet);

                var actual = (implementing as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void ImplementsInterfaceMemberMatchesPublicVariable()
        {
            var intrface =
                @"Option Explicit

Public Foo As Long";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Let TestInterface_Foo(rhs As Long)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyGet);

                var actual = (implementing as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void ImplementsInterfaceMemberLetMatchesPublicIntrinsic()
        {
            var intrface =
                @"Option Explicit

Public Foo As Long";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Let TestInterface_Foo(rhs As Long)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyLet);

                var actual = (implementing as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void ImplementsInterfaceMemberSetDoesNotMatchPublicIntrinsic()
        {
            var intrface =
                @"Option Explicit

Public Foo As Long";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Long
End Property

Public Property Set TestInterface_Foo(rhs As Variant)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertySet);

                Assert.IsFalse((implementing as ModuleBodyElementDeclaration)?.IsInterfaceImplementation);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void ImplementsInterfaceMemberSetMatchesPublicObject()
        {
            var intrface =
                @"Option Explicit

Public Foo As Object";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Object
End Property

Public Property Set TestInterface_Foo(rhs As Object)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertySet);

                var actual = (implementing as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actual, "Expected {0}, resolved to {1}", expected, actual);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void ImplementsInterfaceMemberLetDoesNotMatchPublicObject()
        {
            var intrface =
                @"Option Explicit

Public Foo As Object";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Object
End Property

Public Property Let TestInterface_Foo(rhs As Variant)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var implementing = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyLet);

                Assert.IsFalse((implementing as ModuleBodyElementDeclaration)?.IsInterfaceImplementation);
            }
        }

        [Test]
        [Category("Resolver")]
        [Category("Interfaces")]
        public void ImplementsInterfaceMemberVariantMatchesLetAndSet()
        {
            var intrface =
                @"Option Explicit

Public Foo As Variant";

            var implementation =
                @"Option Explicit

Implements TestInterface

Public Property Get TestInterface_Foo() As Variant
End Property

Public Property Let TestInterface_Foo(rhs As Variant)
End Property

Public Property Set TestInterface_Foo(rhs As Variant)
End Property
";
            var vbe = new MockVbeBuilder()
                .ProjectBuilder("UnderTest", ProjectProtection.Unprotected)
                .AddComponent("TestInterface", ComponentType.ClassModule, intrface, new Selection(1, 1))
                .AddComponent("TestImplementation", ComponentType.ClassModule, implementation, new Selection(1, 1))
                .AddProjectToVbeBuilder()
                .Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var declarations = state.DeclarationFinder.AllDeclarations.ToList();
                var expected = declarations.Single(decl => decl.IdentifierName.Equals("Foo") && decl.DeclarationType == DeclarationType.Variable);
                var setter = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertySet);
                var letter = declarations.Single(decl => decl.IdentifierName.Equals("TestInterface_Foo") && decl.DeclarationType == DeclarationType.PropertyLet);

                var actualSetter = (setter as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;
                var actualLetter = (letter as ModuleBodyElementDeclaration)?.InterfaceMemberImplemented;

                Assert.AreEqual(expected, actualSetter, "Expected {0}, resolved to {1}", expected, actualSetter);
                Assert.AreEqual(expected, actualLetter, "Expected {0}, resolved to {1}", expected, actualLetter);
            }
        }

        private static ClassModuleDeclaration GetTestClassModule(Declaration projectDeclatation, string name, bool isExposed = false)
        {
            var qualifiedClassModuleMemberName = new QualifiedMemberName(StubQualifiedModuleName(name), name);
            var classModuleAttributes = new Attributes();
            if (isExposed)
            {
                classModuleAttributes.AddExposedClassAttribute();
            }
            return new ClassModuleDeclaration(qualifiedClassModuleMemberName, projectDeclatation, name, true, null, classModuleAttributes);
        }

        private static ProjectDeclaration GetTestProject(string name)
        {
            var qualifiedProjectName = new QualifiedMemberName(StubQualifiedModuleName("proj"), name);
            return new ProjectDeclaration(qualifiedProjectName, name, true);
        }

        private static QualifiedModuleName StubQualifiedModuleName(string name)
        {
            return new QualifiedModuleName("dummy", "dummy", name);
        }

        private static FunctionDeclaration GetTestFunction(Declaration moduleDeclatation, string name, Accessibility functionAccessibility)
        {
            var qualifiedFunctionMemberName = new QualifiedMemberName(moduleDeclatation.QualifiedName.QualifiedModuleName, name);
            return new FunctionDeclaration(qualifiedFunctionMemberName, moduleDeclatation, moduleDeclatation, "test", null, "test", functionAccessibility, null, null, Selection.Home, false, true, null, null);
        }

        private static void AddReference(Declaration toDeclaration, Declaration fromModuleDeclaration, ParserRuleContext context = null)
        {
            toDeclaration.AddReference(toDeclaration.QualifiedName.QualifiedModuleName, fromModuleDeclaration, fromModuleDeclaration, context, toDeclaration.IdentifierName, toDeclaration, Selection.Home, new List<Rubberduck.Parsing.Annotations.IParseTreeAnnotation>());
        }
    }
}
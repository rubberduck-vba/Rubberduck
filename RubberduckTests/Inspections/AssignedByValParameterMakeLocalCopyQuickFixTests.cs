using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;
using System.Threading;
using Rubberduck.UI.Refactorings;
using System.Windows.Forms;
using Rubberduck.VBEditor.SafeComWrappers;

namespace RubberduckTests.Inspections
{
    // todo: streamline test cases, currently testing too many things at once, all tests break if newname strategy changes
    [TestClass]
    public class AssignedByValParameterMakeLocalCopyQuickFixTests
    {

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment()
        {
            const string inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            const string expectedCode =
@"Public Sub Foo(ByVal arg1 As String)
Dim localArg1 As String
localArg1 = arg1
    Let localArg1 = ""test""
End Sub";

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUse()
        {
            //Punt if the user-defined is already used in the method
            //string userEnteredName = "userInput";

            string inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub"
;

            var expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
Dim localArg12 As String
localArg12 = arg1
    localArg1 = 6
    Let localArg12 = ""test""
End Sub"
;

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //Modify the suggestion if the auto-generated name is already used in the method
            inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Dim fooVar, localArg1 As Long
    Let arg1 = ""test""
End Sub"
;

            expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
Dim localArg12 As String
localArg12 = arg1
    Dim fooVar, localArg1 As Long
    Let localArg12 = ""test""
End Sub"
;

            quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //handles line continuations
            inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Dim fooVar, _
        localArg1 As Long
    Let arg1 = ""test""
End Sub"
            ;

            expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
Dim localArg12 As String
localArg12 = arg1
    Dim fooVar, _
        localArg1 As Long
    Let localArg12 = ""test""
End Sub"
;
            quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //Punt if the user-defined or auto-generated name is already present as an parameter name
            var userEnteredName = "theSecondArg";

            inputCode =
@"
Public Sub Foo(ByVal arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            expectedCode =
@"
Public Sub Foo(ByVal arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode, userEnteredName);
            Assert.AreEqual(expectedCode, quickFixResult);


        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameDeclaredInSameModule()
        {
            //Punt if the user-defined is already present as a module level declaration
            var userEnteredName = "moduleScopeName2";

            var inputCode =
@"
Private moduleScopeName As String
Private moduleScopeName2 As Long


Public Sub Foo(ByVal arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
    moduleScopeName = arg1 & ""Foo""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            var expectedCode = inputCode;

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode, userEnteredName);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_FunctionNameDeclaredInSameModule()
        {
            //var userEnteredName = "FooFight";

            var inputCode =
@"
Private moduleScopeName As String
Private moduleScopeName2 As Long


Public Sub Foo(ByVal arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
    moduleScopeName = arg1 & ""Foo""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub"
;

            var expectedCode = inputCode;
            string[] invalidNames = { "FooFight", "moduleScopeName", "moduleScopeName2", "Foo" };

            foreach(var name in invalidNames)
            {
                var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode, name);
                Assert.AreEqual(expectedCode, quickFixResult);
            }

        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameAlreadyAccessibleWithinClass()
        {

            var firstClassBody =
@"
Private memberString As String
Private memberLong As Long

Private Sub Class_Initialize()
    memberLong = 6
    memberString = ""No Value""
End Sub

Public Sub Foo(ByVal arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
End Sub

Private Sub FooFight(ByRef arg1 As String)
    xArg1 = 6
    Let arg1 = ""test""
End Sub
";

            var firstClass = new KeyValuePair<string, string>("CFirstClass", firstClassBody);

            var expectedCode = firstClass.Value;
            KeyValuePair<string, string>[] standardModules = { };
            KeyValuePair<string, string>[] classModules = { firstClass };

            //Not changes to the code module if any of these invalid names are chosen
            string[] invalidNames = { "memberLong", "memberString", "FooFight", "Foo" };
            foreach (var invalidName in invalidNames)
            {
                var quickFixResult = GetQuickFixResult(invalidName, firstClass, standardModules, classModules);
                Assert.AreEqual(expectedCode, quickFixResult);
            }

            string[] validNames = { "newVar", "memberString2", "FooFighter", "Food" };
            foreach (var validName in validNames)
            {
                var quickFixResult = GetQuickFixResult(validName, firstClass, standardModules, classModules);
                Assert.AreNotEqual(expectedCode, quickFixResult);
            }
        }

        //Validates which names that are in-scope, and therefore unavailable to be chosen as a variable name
        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameAlreadyAccessibleToProcedure()
        {
            var firstClassBody = 
@"
Private memberString As String
Private memberLong As Long
Private myEggo as String

Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
    memberString = arg1 & ""Foo""
End Sub

Public Function KungFooFighting(ByRef arg1 As String, theSecondArg As Long) As String
    Let arg1 = ""test""
    Dim result As String
    result = arg1 & theSecondArg
    KungFooFighting = result
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
            var secondClassBody = 
@"
Private memberString As String
Private memberLong As Long
Public myEggo as String


Public Sub Foo2( arg1 As String, theSecondArg As Long)
    Let arg1 = ""test""
    memberString = arg1 & ""Foo""
End Sub

Public Function KungFooFighting(ByRef arg1 As String, theSecondArg As Long) As String
    Let arg1 = ""test""
    Dim result As String
    result = arg1 & theSecondArg
    KungFooFighting = result
End Function

Property Let GoMyOtherEggo(newValue As String)
    myEggo = newValue
End Property

Property Get GoMyOtherEggo()
    GoMyOtherEggo = myEggo
End Property

Private Sub FooFighters(ByRef arg1 As String)
    xArg1 = 6
    Let arg1 = ""test""
End Sub
";

            var firstModuleBody =
@"
Option Explicit


Public Const CantTouchThis As String = ""Can't Touch this""
Public THE_FILENAME As String

Sub SetFilename(filename As String)
    THE_FILENAME = filename
End Sub
";

            var secondModuleBody =
@"
Option Explicit


Public BigNumber as Long
Public ShortStory As String

Public Sub DoSomething(filename As String)
    ShortStory = filename
End Sub
";

            var firstClass = new KeyValuePair<string, string>("CFirstClass", firstClassBody);
            var secondClass = new KeyValuePair<string, string>("CSecondClass", secondClassBody);
            var expectedCode = firstClass.Value;

            var firstModule = new KeyValuePair<string, string>("modFirst", firstModuleBody);
            var secondModule = new KeyValuePair<string, string>("modSecond", secondModuleBody);

            KeyValuePair<string, string>[] standardModules = { firstModule, secondModule };
            KeyValuePair<string, string>[] classModules = { firstClass, secondClass };


            //This list of names results are invalid to use - and results in no change to the code module
            string[] inValidNames = { "CantTouchThis", "BigNumber", "DoSomething"
                    , "myEggo", "SetFilename", firstClass.Key, secondClass.Key
                    , firstModule.Key, secondModule.Key};
            foreach (var invalidName in inValidNames)
            {
                var quickFixResult = GetQuickFixResult(invalidName, firstClass, standardModules, classModules);
                Assert.AreEqual(expectedCode, quickFixResult);
            }

            //This list of names results in modifying the code module
            string[] validNames = { "myNewVariable", "Foo2" };
            foreach (var validName in validNames)
            {
                var quickFixResult = GetQuickFixResult(validName, firstClass, standardModules, classModules);
                Assert.AreNotEqual(expectedCode, quickFixResult);
            }
        }

        private string GetQuickFixResult( string userEnteredNames, KeyValuePair<string, string>  resultsComponent, KeyValuePair<string, string>[] standardModules, KeyValuePair<string, string>[] classModules)
        {
            var vbe = BuildProject("TestProject", standardModules.ToList(), classModules.ToList());
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var mockDialogFactory = BuildMockDialogFactory(userEnteredNames);
            var inspection = new AssignedByValParameterInspection(parser.State, mockDialogFactory.Object);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is AssignedByValParameterMakeLocalCopyQuickFix).Fix();

            return GetModuleContent(vbe.Object, resultsComponent.Key);
        }

        private Mock<IVBE> BuildProject( string projectName, List<KeyValuePair<string, string>> moduleComponents, List<KeyValuePair<string,string>> classComponents)
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected);

            moduleComponents.ForEach(c => enclosingProjectBuilder.AddComponent(c.Key, ComponentType.StandardModule, c.Value));
            classComponents.ForEach(c => enclosingProjectBuilder.AddComponent(c.Key, ComponentType.ClassModule, c.Value));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            return builder.Build();
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUseOtherSub()
        {
            //Make sure the modified code stays within the specific method under repair
            const string inputCode =
@"
Public Function Bar2(ByVal arg2 As String) As String
    Dim arg1 As String
    Let arg1 = ""Test1""
    Bar2 = arg1
End Function

Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub

Public Sub Bar(ByVal arg2 As String)
    Dim arg1 As String
    Let arg1 = ""Test2""
End Sub"
;

            const string expectedCode =
@"
Public Function Bar2(ByVal arg2 As String) As String
    Dim arg1 As String
    Let arg1 = ""Test1""
    Bar2 = arg1
End Function

Public Sub Foo(ByVal arg1 As String)
Dim localArg1 As String
localArg1 = arg1
    Let localArg1 = ""test""
End Sub

Public Sub Bar(ByVal arg2 As String)
    Dim arg1 As String
    Let arg1 = ""Test2""
End Sub"
;

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUseOtherProperty()
        {
            //Make sure the modified code stays within the specific method under repair
            var inputCode =
@"
Option Explicit
Private mBar as Long
Public Property Let Foo(ByVal bar As Long)
    bar = 42
    bar = bar * 2
    mBar = bar
End Property

Public Property Get Foo() As Long
    Dim bar as Long
    bar = 12
    Foo = mBar
End Property

Public Function bar() As Long
    Dim xBar As String
    bar = 42
End Function
";
            var expectedCode =
@"
Option Explicit
Private mBar as Long
Public Property Let Foo(ByVal bar As Long)
Dim localBar As Long
localBar = bar
    localBar = 42
    localBar = localBar * 2
    mBar = localBar
End Property

Public Property Get Foo() As Long
    Dim bar as Long
    bar = 12
    Foo = mBar
End Property

Public Function bar() As Long
    Dim xBar As String
    bar = 42
End Function
";

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //variable name use checks do not 'leak' into adjacent procedure(s)
            inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
Dim localArg1 As String
localArg1 = arg1
    Let localArg1 = ""test""
End Sub

Public Sub FooFight(ByRef arg1 As String)
    localArg1 = 6
    Let arg1 = ""test""
End Sub
"
;

            quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_SimilarNamesIgnored()
        {
            //Make sure the modified code stays within the specific method under repair
            var inputCode =
@"
Option Explicit

Public Sub Foo(ByVal bar As Long)
    bar = 42
    bar = bar * 2
    Dim barBell as Long
    barBell = 6
    Dim isobar as Long
    isobar = 13
    Dim bar_candy as Long
    Dim candy_bar as Long
    Dim bar_after_bar as Long
    Dim barsAlot as string
    barsAlot = ""barsAlot:"" & CStr(isobar) & CStr(bar) & CStr(barBell)
    barsAlot = ""barsAlot:"" & CStr(isobar) & CStr( _
        bar) & CStr(barBell)
    total = bar + isobar + candy_bar + bar + bar_candy + barBell + _
            bar_after_bar + bar
bar = 7
    barsAlot = ""bar""
End Sub
";
            var expectedCode =
@"
Option Explicit

Public Sub Foo(ByVal bar As Long)
Dim localBar As Long
localBar = bar
    localBar = 42
    localBar = localBar * 2
    Dim barBell as Long
    barBell = 6
    Dim isobar as Long
    isobar = 13
    Dim bar_candy as Long
    Dim candy_bar as Long
    Dim bar_after_bar as Long
    Dim barsAlot as string
    barsAlot = ""barsAlot:"" & CStr(isobar) & CStr(localBar) & CStr(barBell)
    barsAlot = ""barsAlot:"" & CStr(isobar) & CStr( _
        localBar) & CStr(barBell)
    total = localBar + isobar + candy_bar + localBar + bar_candy + barBell + _
            bar_after_bar + localBar
localBar = 7
    barsAlot = ""bar""
End Sub
";

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_ProperPlacementOfDeclaration()
        {

            var inputCode =
@"Sub DoSomething(_
    ByVal foo As Long, _
    ByRef _
        bar, _
    ByRef barbecue _
                    )
    foo = 4
    bar = barbecue * _
               bar + foo / barbecue
End Sub
";

            var expectedCode =
@"Sub DoSomething(_
    ByVal foo As Long, _
    ByRef _
        bar, _
    ByRef barbecue _
                    )
Dim localFoo As Long
localFoo = foo
    localFoo = 4
    bar = barbecue * _
               bar + localFoo / barbecue
End Sub
";
            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        private string ApplyLocalVariableQuickFixToCodeFragment(string inputCode, string userEnteredName = "")
        {
            var vbe = BuildMockVBEStandardModuleForCodeFragment(inputCode);

            var mockDialogFactory = BuildMockDialogFactory(userEnteredName);

            var inspectionResults = GetInspectionResults(vbe.Object, mockDialogFactory.Object);
            var result = inspectionResults.FirstOrDefault();
            if (result == null)
            {
                Assert.Inconclusive("Inspection yielded no results.");
            }

            result.QuickFixes.Single(s => s is AssignedByValParameterMakeLocalCopyQuickFix).Fix();

            return GetModuleContent(vbe.Object);
        }

        private Mock<IVBE> BuildMockVBEStandardModuleForCodeFragment(string inputCode)
        {
            var builder = new MockVbeBuilder();
            IVBComponent component;
            return builder.BuildFromSingleStandardModule(inputCode, out component);
        }

        private IEnumerable<Rubberduck.Inspections.Abstract.InspectionResultBase> GetInspectionResults(IVBE vbe, IAssignedByValParameterQuickFixDialogFactory mockDialogFactory)
        {
            var parser = MockParser.Create(vbe, new RubberduckParserState(vbe));
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status >= ParserState.Error) { Assert.Inconclusive("Parser Error"); }

            var inspection = new AssignedByValParameterInspection(parser.State, mockDialogFactory);
            return inspection.GetInspectionResults();
        }

        private Mock<IAssignedByValParameterQuickFixDialogFactory> BuildMockDialogFactory(string userEnteredName)
        {
            var mockDialog = new Mock<IAssignedByValParameterQuickFixDialog>();

            mockDialog.SetupAllProperties();

            if (userEnteredName.Length > 0)
            {
                mockDialog.SetupGet(m => m.NewName).Returns(() => userEnteredName);
            }
            mockDialog.SetupGet(m => m.DialogResult).Returns(() => DialogResult.OK);
            
            var mockDialogFactory = new Mock<IAssignedByValParameterQuickFixDialogFactory>();
            mockDialogFactory.Setup(f => f.Create(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<IEnumerable<string>>())).Returns(mockDialog.Object);

            return mockDialogFactory;
        }

        private string GetModuleContent(IVBE vbe, string componentName = "")
        {
            var project = vbe.VBProjects[0];
            var module = componentName.Length >0 
                ?  project.VBComponents[componentName].CodeModule : project.VBComponents[0].CodeModule;
            return module.Content();
        }

        private static RubberduckParserState Parse(Mock<IVBE> vbe)
        {
            var parser = MockParser.Create(vbe.Object, new RubberduckParserState(vbe.Object));
            parser.Parse(new CancellationTokenSource());
            if (parser.State.Status != ParserState.Ready)
            {
                Assert.Inconclusive("Parser state should be 'Ready', but returns '{0}'.", parser.State.Status);
            }
            var state = parser.State;
            return state;
        }
    }
}

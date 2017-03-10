using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections;
using Rubberduck.Inspections.QuickFixes;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.UI.Refactorings;
using System.Windows.Forms;

namespace RubberduckTests.Inspections
{
    [TestClass]
    public class AssignedByValParameterMakeLocalCopyQuickFixTests
    {

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment()
        {
            var inputCode =
@"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            var expectedCode =
@"Public Sub Foo(ByVal arg1 As String)
Dim localArg1 As String
localArg1 = arg1
    Let localArg1 = ""test""
End Sub";

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

            //weaponized formatting
            inputCode =
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

            expectedCode =
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
            quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_ComputedNameAvoidsCollision()
        {
            var inputCode =
@"
Public Sub Foo(ByVal arg1 As String)
    Dim fooVar, _
        localArg1 As Long
    Let arg1 = ""test""
End Sub"
;

            var expectedCode =
@"
Public Sub Foo(ByVal arg1 As String)
Dim localArg12 As String
localArg12 = arg1
    Dim fooVar, _
        localArg1 As Long
    Let localArg12 = ""test""
End Sub"
;
            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);

        }

        [TestMethod]
        [TestCategory("Inspections")]
        public void AssignedByValParameter_LocalVariableAssignment_NameInUseOtherSub()
        {
            //Make sure the modified code stays within the specific method under repair
            var inputCode =
@"
Public Function Bar2(ByVal arg2 As String) As String
    Dim arg1 As String
    Let arg1 = ""Test1""
    Bar2 = arg1
End Function

Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub

'VerifyNoChangeBelowThisLine
Public Sub Bar(ByVal arg2 As String)
    Dim arg1 As String
    Let arg1 = ""Test2""
End Sub"
;
            var expectedFragment =
@"
Public Sub Bar(ByVal arg2 As String)
    Dim arg1 As String
    Let arg1 = ""Test2""
End Sub"
;
            string[] splitToken = { "'VerifyNoChangeBelowThisLine" };
            var expectedCode = inputCode.Split(splitToken, System.StringSplitOptions.None)[1];

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            var evaluatedFragment = quickFixResult.Split(splitToken, System.StringSplitOptions.None)[1];
            Assert.AreEqual(expectedFragment, evaluatedFragment);
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

'VerifyNoChangeBelowThisLine
Public Function bar() As Long
    Dim localBar As Long
    localBar = 7
    bar = localBar
End Function
";
            string[] splitToken = { "'VerifyNoChangeBelowThisLine" };
            var expectedCode = inputCode.Split(splitToken, System.StringSplitOptions.None)[1];

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            var evaluatedResult = quickFixResult.Split(splitToken, System.StringSplitOptions.None)[1];
            Assert.AreEqual(expectedCode, evaluatedResult);

        }

        private string ApplyLocalVariableQuickFixToCodeFragment(string inputCode, string userEnteredName = "")
        {
            var vbe = BuildMockVBE(inputCode);

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

        private Mock<IVBE> BuildMockVBE(string inputCode)
        {
            IVBComponent component;
            return MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
        }

        private IEnumerable<Rubberduck.Inspections.Abstract.InspectionResultBase> GetInspectionResults(IVBE vbe, IAssignedByValParameterQuickFixDialogFactory mockDialogFactory)
        {
            var state = MockParser.CreateAndParse(vbe);

            var inspection = new AssignedByValParameterInspection(state, mockDialogFactory);
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
            var module = componentName.Length > 0
                ? project.VBComponents[componentName].CodeModule : project.VBComponents[0].CodeModule;
            return module.Content();
        }
/*//<<<<<<< HEAD
//=======

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

        private string GetQuickFixResult(string userEnteredNames, TestComponentSpecification resultsComponent, TestComponentSpecification[] testComponents)
        {
            var vbe = BuildProject("TestProject", testComponents.ToList());
            var state = MockParser.CreateAndParse(vbe.Object);

            var mockDialogFactory = BuildMockDialogFactory(userEnteredNames);
            var inspection = new AssignedByValParameterInspection(state, mockDialogFactory.Object);
            var inspectionResults = inspection.GetInspectionResults();

            inspectionResults.First().QuickFixes.Single(s => s is AssignedByValParameterMakeLocalCopyQuickFix).Fix();

            return GetModuleContent(vbe.Object, resultsComponent.Name);
        }

        private Mock<IVBE> BuildProject(string projectName, List<TestComponentSpecification> testComponents)
        {
            var builder = new MockVbeBuilder();
            var enclosingProjectBuilder = builder.ProjectBuilder(projectName, ProjectProtection.Unprotected);

            testComponents.ForEach(c => enclosingProjectBuilder.AddComponent(c.Name, c.ModuleType, c.Content));
            var enclosingProject = enclosingProjectBuilder.Build();
            builder.AddProject(enclosingProject);
            return builder.Build();
        }

        private string GetNameAlreadyAccessibleWithinClass_FirstClassBody()
        {
            return
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

        }
        private string GetRespectsDeclarationsAccessibilityRules_FirstClassBody()
        {
            return
@"
Private memberString As String
Private memberLong As Long
Private myEggo as String

Public Sub Foo(ByVal arg1 As String)
    Dim localVar as Long
    localVar = 7
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
        }
        private string GetRespectsDeclarationsAccessibilityRules_SecondClassBody()
        {
            return
@"
Private memberString As String
Private memberLong As Long
Public mySecondEggo as String


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
    result = KungFooFighting(st, v)
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
//>>>>>>> rubberduck-vba/next*/
    }
}

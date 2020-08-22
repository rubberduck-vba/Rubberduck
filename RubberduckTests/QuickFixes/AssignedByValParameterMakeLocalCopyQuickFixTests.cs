using NUnit.Framework;
using System.Linq;
using System.Threading;
using Moq;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;
using Rubberduck.UI.Refactorings;
using System.Windows.Forms;
using System;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class AssignedByValParameterMakeLocalCopyQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_LocalVariableAssignment_Sub()
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
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_LocalVariableAssignment_SubWithCommentn()
        {
            var inputCode =
@"Public Sub Foo(ByVal arg1 As String)  'Comment is here _
and here _
and here
    Let arg1 = ""test""
End Sub";
            var expectedCode =
@"Public Sub Foo(ByVal arg1 As String)  'Comment is here _
and here _
and here
    Dim localArg1 As String
    localArg1 = arg1
    Let localArg1 = ""test""
End Sub";

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_LocalVariableAssignment_Function()
        {
            var inputCode =
@"Public Function Foo(ByVal arg1 As String) As String
    arg1 = ""test""
    Foo = arg1
End Function";
            var expectedCode =
@"Public Function Foo(ByVal arg1 As String) As String
    Dim localArg1 As String
    localArg1 = arg1
    localArg1 = ""test""
    Foo = localArg1
End Function";

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_LocalVariableAssignment_FunctionWithComment()
        {
            var inputCode =
@"Public Function Foo(ByVal arg1 As String) As String   'This is a comment
    arg1 = ""test""
    Foo = arg1
End Function";
            var expectedCode =
@"Public Function Foo(ByVal arg1 As String) As String   'This is a comment
    Dim localArg1 As String
    localArg1 = arg1
    localArg1 = ""test""
    Foo = localArg1
End Function";

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_LocalVariableAssignment_Property()
        {
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
";
            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }
        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_LocalVariableAssignment_ComplexFormat()
        {
            var inputCode =
                @"Sub DoSomething( _
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
                @"Sub DoSomething( _
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

        [Test]
        [Category("QuickFixes")]
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

        [Test]
        [Category("QuickFixes")]
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
            string[] splitToken = { "'VerifyNoChangeBelowThisLine" };
            var expectedCode = inputCode.Split(splitToken, System.StringSplitOptions.None)[1];

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            var evaluatedResult = quickFixResult.Split(splitToken, System.StringSplitOptions.None)[1];

            Assert.AreEqual(expectedCode, evaluatedResult);
        }

        [Test]
        [Category("QuickFixes")]
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

        //Replicates issue #2873 : AssignedByValParameter quick fix needs to use `Set` for reference types.
        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_LocalVariableAssignment_UsesSet()
        {
            var inputCode =
                    @"
Public Sub Foo(ByVal target As Range)
    Set target = Selection
End Sub"
                ;
            var expectedCode =
                    @"
Public Sub Foo(ByVal target As Range)
    Dim localTarget As Range
    Set localTarget = target
    Set localTarget = Selection
End Sub"
                ;

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_LocalVariableAssignment_NoAsTypeClause()
        {
            var inputCode =
                    @"
Public Sub Foo(FirstArg As Long, ByVal arg1)
    arg1 = Range(""A1: C4"")
End Sub"
                ;
            var expectedCode =
                    @"
Public Sub Foo(FirstArg As Long, ByVal arg1)
    Dim localArg1 As Variant
    localArg1 = arg1
    localArg1 = Range(""A1: C4"")
End Sub"
                ;

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_LocalVariableAssignment_EnumType()
        {
            var inputCode =
                    @"
Enum TestEnum
    EnumOne
    EnumTwo
    EnumThree
End Enum

Public Sub Foo(FirstArg As Long, ByVal arg1 As TestEnum)
    arg1 = EnumThree
End Sub"
                ;
            var expectedCode =
                    @"
Enum TestEnum
    EnumOne
    EnumTwo
    EnumThree
End Enum

Public Sub Foo(FirstArg As Long, ByVal arg1 As TestEnum)
    Dim localArg1 As TestEnum
    localArg1 = arg1
    localArg1 = EnumThree
End Sub"
                ;

            var quickFixResult = ApplyLocalVariableQuickFixToCodeFragment(inputCode);
            Assert.AreEqual(expectedCode, quickFixResult);
        }

        private string ApplyLocalVariableQuickFixToCodeFragment(string inputCode, string userEnteredName = "")
        {
            var vbe = BuildMockVBE(inputCode, out var component);

            var mockDialogFactory = BuildMockDialogFactory(userEnteredName);

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var inspection = new AssignedByValParameterInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
                var result = inspectionResults.FirstOrDefault();
                if (result == null)
                {
                    Assert.Inconclusive("Inspection yielded no results.");
                }

                var rewriteSession = rewritingManager.CheckOutCodePaneSession();

                new AssignedByValParameterMakeLocalCopyQuickFix(state, mockDialogFactory.Object).Fix(result, rewriteSession);

                return rewriteSession.CheckOutModuleRewriter(component.QualifiedModuleName).GetText();
            }
        }

        private Mock<IVBE> BuildMockVBE(string inputCode, out IVBComponent component)
        {
            return MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out component);
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
            mockDialogFactory.Setup(f => f.Create(It.IsAny<string>(), It.IsAny<string>(), It.IsAny<Func<string,bool>>())).Returns(mockDialog.Object);

            return mockDialogFactory;
        }
    }
}

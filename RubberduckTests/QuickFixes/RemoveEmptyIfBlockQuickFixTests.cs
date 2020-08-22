using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveEmptyIfBlockQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesLoneIf()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesSingleLineIf()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then Else Bar
End Sub

Sub Bar()
End Sub";

            const string expectedCode =
                @"Sub Foo()
    If Not True Then Bar
End Sub

Sub Bar()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesLoneIf_WithComment()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        ' Im a comment
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesIf_WithElseIfAndElse()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    ElseIf False Then
        Dim d
        d = 0
    Else
        Dim b
        b = 0
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If False Then
        Dim d
        d = 0
    Else
        Dim b
        b = 0
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElseIf()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Dim d
        d = 0
    ElseIf False Then
    Else
        Dim b
        b = 0
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True Then
        Dim d
        d = 0
    Else
        Dim b
        b = 0
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElseIf_HasComment()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Dim d
        d = 0
    ElseIf False Then
        ' Im a comment
    Else
        Dim b
        b = 0
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True Then
        Dim d
        d = 0
    Else
        Dim b
        b = 0
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesIf_HasVariable()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Dim d
    Else
        Dim b
        b = 0
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    Dim d
    If Not True Then
        
        Dim b
        b = 0
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesIf_HasVariable_WithComment()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        ' comment
        Dim d
    Else
        Dim b
        b = 0
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    Dim d
    If Not True Then
        ' comment
        
        Dim b
        b = 0
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesIf_HasVariable_WithLabel()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
5       Dim d
a:      Dim e
15 b:   Dim f
    Else
        Dim b
        b = 0
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    
5       Dim d
a:      Dim e
15 b:   Dim f
    If Not True Then

        Dim b
        b = 0
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesIf_HasConst()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Const d = 0
    Else
        Dim b
        b = 0
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    Const d = 0
    If Not True Then
        
        Dim b
        b = 0
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElseIf_HasVariable()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Dim d
        d = 0
    ElseIf True Then
        Dim b
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    Dim b
    If True Then
        Dim d
        d = 0
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElseIf_HasConst()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
        Dim b
        b = 0
    ElseIf True Then
        Const d = 0
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    Const d = 0
    If True Then
        Dim b
        b = 0
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesIf_UpdatesElseIf()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    ElseIf False Then
        Dim d
        d = 0
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If False Then
        Dim d
        d = 0
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_SimpleCondition()
        {
            const string inputCode =
                @"Sub Foo()
    If True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If Not True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_Equals()
        {
            const string inputCode =
                @"Sub Foo()
    If True = True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True <> True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_NotEquals()
        {
            const string inputCode =
                @"Sub Foo()
    If True <> True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True = True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_LessThan()
        {
            const string inputCode =
                @"Sub Foo()
    If True < True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True >= True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_LessThanEquals()
        {
            const string inputCode =
                @"Sub Foo()
    If True <= True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True > True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_GreaterThan()
        {
            const string inputCode =
                @"Sub Foo()
    If True > True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True <= True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_GreaterThanEquals()
        {
            const string inputCode =
                @"Sub Foo()
    If True >= True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True < True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_Not()
        {
            const string inputCode =
                @"Sub Foo()
    If Not True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_Not_NoWhitespace()
        {
            const string inputCode =
                @"Sub Foo()
    If Not(True) Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If (True) Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_And()
        {
            const string inputCode =
                @"Sub Foo()
    If True And True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True Or True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_Or()
        {
            const string inputCode =
                @"Sub Foo()
    If True Or True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True And True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_Xor()
        {
            const string inputCode =
                @"Sub Foo()
    If True Xor True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If Not (True Xor True) Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_ComplexCondition()
        {
            const string inputCode =
                @"Sub Foo()
    If True Or True And True Or True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True Or True And True And True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_ComplexCondition1()
        {
            const string inputCode =
                @"Sub Foo()
    If True And True Or True And True Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If True And True And True And True Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_ComplexCondition_WithParentheses()
        {
            const string inputCode =
                @"Sub Foo()
    If (True Or True) And (True Or True) Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If (True Or True) Or (True Or True) Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyIfBlock_QuickFixRemovesElse_InvertsIf_ComplexCondition2()
        {
            const string inputCode =
                @"Sub Foo()
    If 1 > 2 And 3 = 3 Or 4 <> 5 And 8 - 6 = 2 Then
    Else
    End If
End Sub";
            const string expectedCode =
                @"Sub Foo()
    If 1 > 2 And 3 = 3 And 4 <> 5 And 8 - 6 = 2 Then
    
    End If
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new EmptyIfBlockInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveEmptyIfBlockQuickFix();
        }
    }
}

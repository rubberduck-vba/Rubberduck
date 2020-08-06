using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class IgnoreOnceQuickFixTests
    {
        [Test]
        [Category("QuickFixes")]
        public void ApplicationWorksheetFunction_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub ExcelSub()
    Dim foo As Double
    foo = Application.Pi
End Sub";

            const string expectedCode =
                @"Sub ExcelSub()
    Dim foo As Double
    '@Ignore ApplicationWorksheetFunction
    foo = Application.Pi
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ApplicationWorksheetFunctionInspection(state), TestStandardModuleInExcelVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void AssignedByValParameter_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            const string expectedCode =
                @"'@Ignore AssignedByValParameter
Public Sub Foo(ByVal arg1 As String)
    Let arg1 = ""test""
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new AssignedByValParameterInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ConstantNotUsed_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const const1 As Integer = 9
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    '@Ignore ConstantNotUsed
    Const const1 As Integer = 9
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ConstantNotUsedInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EmptyStringLiteral_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo(ByRef arg1 As String)
    arg1 = """"
End Sub";

            const string expectedCode =
                @"Public Sub Foo(ByRef arg1 As String)
    '@Ignore EmptyStringLiteral
    arg1 = """"
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new EmptyStringLiteralInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void EncapsulatePublicField_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Public fizz As Boolean";

            const string expectedCode =
                @"'@Ignore EncapsulatePublicField
Public fizz As Boolean";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new EncapsulatePublicFieldInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category("Unused Value")]
        public void FunctionReturnValueDiscarded_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    Foo ""test""
End Sub";

            const string expectedCode =
                @"Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    '@Ignore FunctionReturnValueDiscarded
    Foo ""test""
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new FunctionReturnValueDiscardedInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Category("Unused Value")]
        public void FunctionReturnValueAlwaysDiscarded_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    Foo ""test""
End Sub";

            const string expectedCode =
                @"'@Ignore FunctionReturnValueAlwaysDiscarded
Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    Foo ""test""
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new FunctionReturnValueAlwaysDiscardedInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void AnnotationListFollowedByCommentAddsAnnotationCorrectly()
        {
            const string inputCode = @"
Public Function GetSomething() As Long
    '@Ignore VariableNotAssigned: Is followed by a comment.
    Dim foo
    GetSomething = foo
End Function
";

            const string expectedCode = @"
Public Function GetSomething() As Long
    '@Ignore VariableTypeNotDeclared, VariableNotAssigned: Is followed by a comment.
    Dim foo
    GetSomething = foo
End Function
";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new VariableTypeNotDeclaredInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitActiveSheetReference_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub foo()
    Dim arr1() As Variant
    arr1 = Range(""A1:B2"")
End Sub";

            const string expectedCode =
                @"Sub foo()
    Dim arr1() As Variant
    '@Ignore ImplicitActiveSheetReference
    arr1 = Range(""A1:B2"")
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ImplicitActiveSheetReferenceInspection(state), TestClassInExcelVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitActiveWorkbookReference_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    Set sheet = Worksheets(""Sheet1"")
End Sub";

            const string expectedCode =
                @"
Sub foo()
    Dim sheet As Worksheet
    '@Ignore ImplicitActiveWorkbookReference
    Set sheet = Worksheets(""Sheet1"")
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ImplicitActiveWorkbookReferenceInspection(state), TestClassInExcelVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void ImplicitByRefModifier_IgnoredQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo(arg1 As Integer)
End Sub";

            const string expectedCode =
                @"'@Ignore ImplicitByRefModifier
Sub Foo(arg1 As Integer)
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ImplicitByRefModifierInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }



        [Test]
        [Category("QuickFixes")]
        public void ImplicitPublicMember_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo(ByVal arg1 as Integer)
'Just an inoffensive little comment

End Sub";

            const string expectedCode =
                @"'@Ignore ImplicitPublicMember
Sub Foo(ByVal arg1 as Integer)
'Just an inoffensive little comment

End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ImplicitPublicMemberInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitVariantReturnType_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Function Foo()
End Function";

            const string expectedCode =
                @"'@Ignore ImplicitVariantReturnType
Function Foo()
End Function";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ImplicitVariantReturnTypeInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]

        [Category("QuickFixes")]
        public void LabelNotUsed_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
label1:
End Sub";

            const string expectedCode =
                @"Sub Foo()
'@Ignore LineLabelNotUsed
label1:
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new LineLabelNotUsedInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ModuleScopeDimKeyword_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Dim foo";

            const string expectedCode =
                @"'@Ignore ModuleScopeDimKeyword
Dim foo";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ModuleScopeDimKeywordInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void MoveFieldCloserToUsage_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Private bar As String
Public Sub Foo()
    bar = ""test""
End Sub";

            const string expectedCode =
                @"'@Ignore MoveFieldCloserToUsage
Private bar As String
Public Sub Foo()
    bar = ""test""
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new MoveFieldCloserToUsageInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Ignore("With the current annotation scoping rules, this test makes no sense since the Ignore annotation will not attach to the offending context.")]
        public void MultilineParameter_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo( _
    ByVal _
    Var1 _
    As _
    Integer)
End Sub";

            const string expectedCode =
                @"'@Ignore MultilineParameter
Public Sub Foo( _
    ByVal _
    Var1 _
    As _
    Integer)
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new MultilineParameterInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void MultipleDeclarations_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer, var2 As String
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    '@Ignore MultipleDeclarations
    Dim var1 As Integer, var2 As String
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new MultipleDeclarationsInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonReturningFunction_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Function Foo() As Boolean
End Function";

            const string expectedCode =
                @"'@Ignore NonReturningFunction
Function Foo() As Boolean
End Function";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new NonReturningFunctionInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void ObjectVariableNotSet_IgnoreQuickFixWorks()
        {
            var inputCode =
                @"
Private Sub DoSomething()
    
    Dim target As Class1
    target = New Class1
    
    target.Value = ""forgot something?""

End Sub";
            var expectedCode =
                @"
Private Sub DoSomething()
    
    Dim target As Class1
    '@Ignore ObjectVariableNotSet
    target = New Class1
    
    target.Value = ""forgot something?""

End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ObjectVariableNotSetInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCallStatement_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    Call Foo
End Sub";

            const string expectedCode =
                @"Sub Foo()
    '@Ignore ObsoleteCallStatement
    Call Goo(1, ""test"")
End Sub

Sub Goo(arg1 As Integer, arg1 As String)
    '@Ignore ObsoleteCallStatement
    Call Foo
End Sub";

            var actualCode = ApplyIgnoreOnceToAllResults(inputCode, state => new ObsoleteCallStatementInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteCommentSyntax_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Rem test1";

            const string expectedCode =
                @"'@Ignore ObsoleteCommentSyntax
Rem test1";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ObsoleteCommentSyntaxInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteErrorSyntax_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
    Error 91
End Sub";

            const string expectedCode =
                @"Sub Foo()
    '@Ignore ObsoleteErrorSyntax
    Error 91
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ObsoleteErrorSyntaxInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteGlobal_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Global var1 As Integer";

            const string expectedCode =
                @"'@Ignore ObsoleteGlobal
Global var1 As Integer";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ObsoleteGlobalInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void ObsoleteLetStatement_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    Let var2 = var1
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Dim var1 As Integer
    Dim var2 As Integer
    
    '@Ignore ObsoleteLetStatement
    Let var2 = var1
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ObsoleteLetStatementInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Public Function Foo$(ByVal fizz As Integer)
    Foo = ""test""
End Function";

            const string expectedCode =
                @"'@Ignore ObsoleteTypeHint
Public Function Foo$(ByVal fizz As Integer)
    Foo = ""test""
End Function";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ObsoleteTypeHintInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void OptionBaseOneSpecified_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Option Base 1";

            const string expectedCode =
                @"'@Ignore OptionBase
Option Base 1";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new OptionBaseInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void ParameterCanBeByVal_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo(ByRef _
arg1 As String)
End Sub";

            const string expectedCode =
                @"'@Ignore ParameterCanBeByVal
Sub Foo(ByRef _
arg1 As String)
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ParameterCanBeByValInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void GivenPrivateSub_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            const string expectedCode =
                @"'@Ignore ParameterNotUsed
Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ParameterNotUsedInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void ProcedureNotUsed_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            const string expectedCode =
                @"'@Ignore ProcedureNotUsed
Private Sub Foo(ByVal arg1 as Integer)
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ProcedureNotUsedInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef arg1 As Integer)
    arg1 = 42
End Sub";

            const string expectedCode =
                @"'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub Foo(ByRef arg1 As Integer)
    arg1 = 42
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ProcedureCanBeWrittenAsFunctionInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void RedundantByRefModifier_IgnoredQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
                @"'@Ignore RedundantByRefModifier
Sub Foo(ByRef arg1 As Integer)
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new RedundantByRefModifierInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void SelfAssignedDeclaration_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
    Dim b As New Collection
End Sub";

            const string expectedCode =
                @"Sub Foo()
    '@Ignore SelfAssignedDeclaration
    Dim b As New Collection
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new SelfAssignedDeclarationInspection(state), TestClassVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariableUsage_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean
    bb = b
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim b As Boolean
    Dim bb As Boolean
    '@Ignore UnassignedVariableUsage
    bb = b
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new UnassignedVariableUsageInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        [Ignore("Broken feature - passes locally but not in AV - see FIXME below")]
        public void UntypedFunctionUsage_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
    Dim str As String
    str = Left(""test"", 1)
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim str As String
'@Ignore UntypedFunctionUsage
    str = Left(""test"", 1)
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .AddReference(ReferenceLibrary.VBA)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var component = project.Object.VBComponents[0];
            var (parser, rewritingManager) = MockParser.CreateWithRewriteManager(vbe.Object);
            using (var state = parser.State)
            {
                //FIXME reinstate and unignore tests
                //refers to "UntypedFunctionUsageInspectionTests.GetBuiltInDeclarations()"
                //GetBuiltInDeclarations().ForEach(d => state.AddDeclaration(d));

                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new UntypedFunctionUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();

                new IgnoreOnceQuickFix(new AnnotationUpdater(state), state, new[] { inspection }).Fix(inspectionResults.First(), rewriteSession);
                var actualCode = rewriteSession.CheckOutModuleRewriter(component.QualifiedModuleName).GetText();

                Assert.AreEqual(expectedCode, actualCode);
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void UseMeaningfulName_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Ffffff()
End Sub";

            const string expectedCode =
                @"'@Ignore UseMeaningfulName
Sub Ffffff()
End Sub";

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var component = project.Object.VBComponents[0];
            var vbe = builder.AddProject(project).Build();

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {
                var inspection = new UseMeaningfulNameInspection(state, UseMeaningfulNameInspectionTests.GetInspectionSettings().Object);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();

                new IgnoreOnceQuickFix(new AnnotationUpdater(state), state, new[] { inspection }).Fix(inspectionResults.First(), rewriteSession);
                var actualCode = rewriteSession.CheckOutModuleRewriter(component.QualifiedModuleName).GetText();

                Assert.AreEqual(expectedCode, actualCode);
            }
        }


        [Test]
        [Category("QuickFixes")]
        public void UnassignedVariable_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 as Integer
End Sub";

            const string expectedCode =
                @"Sub Foo()
'@Ignore VariableNotAssigned
Dim var1 as Integer
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new VariableNotAssignedInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void UnusedVariable_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
Dim var1 As String
End Sub";

            const string expectedCode =
                @"Sub Foo()
'@Ignore VariableNotUsed
Dim var1 As String
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new VariableNotUsedInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void VariableTypeNotDeclared_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo(arg1)
End Sub";

            const string expectedCode =
                @"'@Ignore VariableTypeNotDeclared
Sub Foo(arg1)
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new VariableTypeNotDeclaredInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void EmptyModule_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Option Explicit";

            const string expectedCode =
                @"'@IgnoreModule EmptyModule
Option Explicit";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new EmptyModuleInspection(state, state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }


        [Test]
        [Category("QuickFixes")]
        public void ModuleWithoutFolder_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Option Explicit";

            const string expectedCode =
                @"'@IgnoreModule ModuleWithoutFolder
Option Explicit";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ModuleWithoutFolderInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void WriteOnlyProperty_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Property Let Foo(value)
End Property";

            const string expectedCode =
                @"'@Ignore WriteOnlyProperty
Property Let Foo(value)
End Property";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new WriteOnlyPropertyInspection(state), TestClassVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void BooleanAssignedInIfElseInspection_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Sub Foo()
    Dim d As Boolean
    If True Then
        d = True
    Else
        d = False
    EndIf
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Dim d As Boolean
    '@Ignore BooleanAssignedInIfElse
    If True Then
        d = True
    Else
        d = False
    EndIf
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new BooleanAssignedInIfElseInspection(state), TestClassVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IgnoreQuickFixAppendsToExistingAnnotation()
        {
            const string inputCode =
                @"Sub Foo()
    '@Ignore VariableNotUsed
    x = 42
End Sub";

            const string expectedCode =
                @"Sub Foo()
    '@Ignore UndeclaredVariable, VariableNotUsed
    x = 42
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new UndeclaredVariableInspection(state), TestClassVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IgnoreQuickFixPrependsToExistingAnnotation_Module()
        {
            const string inputCode =
                @"'@IgnoreModule EmptyModule
Option Explicit";

            const string expectedCode =
                @"'@IgnoreModule ModuleWithoutFolder, EmptyModule
Option Explicit";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ModuleWithoutFolderInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IgnoreQuickFixDoesNotAppendToExistingAnnotationMixed_ModuleAfterNonModule()
        {
            const string inputCode =
                @"'@Ignore ParameterCanBeByVal
Private Sub Foo(arg)
End Sub";

            const string expectedCode =
                @"'@IgnoreModule ModuleWithoutFolder
'@Ignore ParameterCanBeByVal
Private Sub Foo(arg)
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ModuleWithoutFolderInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IgnoreQuickFixDoesNotPrependToExistingAnnotationMixed_NonModuleAfterModule()
        {
            const string inputCode =
                @"'@IgnoreModule ModuleWithoutFolder
Private Sub Foo(arg)
End Sub";

            const string expectedCode =
                @"'@IgnoreModule ModuleWithoutFolder
'@Ignore ParameterCanBeByVal
Private Sub Foo(arg)
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new ParameterCanBeByValInspection(state), TestStandardModuleVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IgnoreQuickFixAddsAnnotationAfterComment()
        {
            const string inputCode =
                @"Sub Foo()
    'comment
    x = 42
End Sub";

            const string expectedCode =
                @"Sub Foo()
    'comment
    '@Ignore UndeclaredVariable
    x = 42
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new UndeclaredVariableInspection(state), TestClassVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IgnoreQuickFixAddsAnnotationAfterRemComment()
        {
            const string inputCode =
                @"Sub Foo()
    Rem comment
    x = 42
End Sub";

            const string expectedCode =
                @"Sub Foo()
    Rem comment
    '@Ignore UndeclaredVariable
    x = 42
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new UndeclaredVariableInspection(state), TestClassVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void IgnoreQuickFixAddsAnnotationAfterMultilineComment()
        {
            const string inputCode =
                @"Sub Foo()
    'multi _
     line _
     comment
    x = 42
End Sub";

            const string expectedCode =
                @"Sub Foo()
    'multi _
     line _
     comment
    '@Ignore UndeclaredVariable
    x = 42
End Sub";

            var actualCode = ApplyIgnoreOnceToFirstResult(inputCode, state => new UndeclaredVariableInspection(state), TestClassVbeSetup);
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ImplicitlyTypedConst_IgnoreOnceQuickFixWorks()
        {
            const string inputCode =
@"Public Sub Foo()
    Const bar = 0
End Sub";
            
            const string expected =
@"Public Sub Foo()
    '@Ignore ImplicitlyTypedConst
    Const bar = 0
End Sub";

            var actual = ApplyIgnoreOnceToFirstResult(inputCode, state => new ImplicitlyTypedConstInspection(state), TestStandardModuleVbeSetup);

            Assert.AreEqual(expected, actual);
        }

        private string ApplyIgnoreOnceToFirstResult(
            string inputCode,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            Func<string, (IVBE vbe, QualifiedModuleName moduleName)> vbeSetup)
        {
            var (vbe, moduleName) = vbeSetup(inputCode);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var inspection = inspectionFactory(state);
                var inspectionResults = InspectionResults(inspection, state);
                var resultToFix = inspectionResults.First();
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();

                var quickFix = new IgnoreOnceQuickFix(new AnnotationUpdater(state), state, new[] {inspection});
                quickFix.Fix(resultToFix, rewriteSession);

                return rewriteSession.CheckOutModuleRewriter(moduleName).GetText();
            }
        }

        private IEnumerable<IInspectionResult> InspectionResults(IInspection inspection, RubberduckParserState state)
        {
            if (inspection is IParseTreeInspection)
            {
                var inspector = InspectionsHelper.GetInspector(inspection);
                return inspector.FindIssuesAsync(state, CancellationToken.None).Result;
            }

            return inspection.GetInspectionResults(CancellationToken.None);
        }

        private (IVBE vbe, QualifiedModuleName moduleName) TestClassVbeSetup(string inputCode)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var component = project.Object.VBComponents[0];
            var vbe = builder.AddProject(project).Build();

            return (vbe.Object, component.QualifiedModuleName);
        }

        private (IVBE vbe, QualifiedModuleName moduleName) TestStandardModuleVbeSetup(string inputCode)
        {
            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            return (vbe.Object, component.QualifiedModuleName);
        }

        private (IVBE vbe, QualifiedModuleName moduleName) TestClassInExcelVbeSetup(string inputCode)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode)
                .AddReference(ReferenceLibrary.Excel)
                .Build();
            var component = project.Object.VBComponents[0];
            var vbe = builder.AddProject(project).Build();

            return (vbe.Object, component.QualifiedModuleName);
        }

        private (IVBE vbe, QualifiedModuleName moduleName) TestStandardModuleInExcelVbeSetup(string inputCode)
        {
            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference(ReferenceLibrary.Excel)
                .Build();

            var vbe = builder.AddProject(project).Build();
            var component = vbe.Object.SelectedVBComponent;

            return (vbe.Object, component.QualifiedModuleName);
        }

        private string ApplyIgnoreOnceToAllResults(
            string inputCode,
            Func<RubberduckParserState, IInspection> inspectionFactory,
            Func<string, (IVBE vbe, QualifiedModuleName moduleName)> vbeSetup)
        {
            var (vbe, moduleName) = vbeSetup(inputCode);
            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe);
            using (state)
            {
                var inspection = inspectionFactory(state);
                var inspectionResults = InspectionResults(inspection, state);
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();

                var quickFix = new IgnoreOnceQuickFix(new AnnotationUpdater(state), state, new[] { inspection });

                foreach (var resultToFix in inspectionResults)
                {
                    quickFix.Fix(resultToFix, rewriteSession);
                }

                return rewriteSession.CheckOutModuleRewriter(moduleName).GetText();
            }
        }
    }
}

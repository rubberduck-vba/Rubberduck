using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;
using RubberduckTests.Inspections;
using Rubberduck.Parsing.Inspections.Resources;
using RubberduckTests.Common;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class IgnoreOnceQuickFixTests
    {
        [Test]
        [DeploymentItem(@"Testfiles\")]
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("Module1", ComponentType.StandardModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();

            var vbe = builder.AddProject(project).Build();
            var component = vbe.Object.SelectedVBComponent;

            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                state.AddTestLibrary("Excel.1.8.xml");

                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new ApplicationWorksheetFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new AssignedByValParameterInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ConstantNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EmptyStringLiteralInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new EncapsulatePublicFieldInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        [Category("Unused Value")]
        public void FunctionReturnValueNotUsed_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    Foo ""test""
End Sub";

            const string expectedCode =
                @"'@Ignore FunctionReturnValueNotUsed
Public Function Foo(ByVal bar As String) As Boolean
End Function

Public Sub Goo()
    Foo ""test""
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new FunctionReturnValueNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [DeploymentItem(@"TestFiles\")]
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();
            var component = project.Object.VBComponents[0];
            var vbe = builder.AddProject(project).Build();

            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                state.AddTestLibrary("Excel.1.8.xml");

                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new ImplicitActiveSheetReferenceInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [DeploymentItem(@"TestFiles\")]
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("TestProject1", "TestProject1", ProjectProtection.Unprotected)
                .AddComponent("Class1", ComponentType.ClassModule, inputCode)
                .AddReference("Excel", MockVbeBuilder.LibraryPathMsExcel, 1, 8, true)
                .Build();
            var component = project.Object.VBComponents[0];
            var vbe = builder.AddProject(project).Build();


            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                state.AddTestLibrary("Excel.1.8.xml");

                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new ImplicitActiveWorkbookReferenceInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitByRefModifierInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitPublicMemberInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ImplicitVariantReturnTypeInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new LineLabelNotUsedInspection(state);
                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ModuleScopeDimKeywordInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MoveFieldCloserToUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MultilineParameterInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new MultipleDeclarationsInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new NonReturningFunctionInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }


        [Test]
        [Category("QuickFixes")]
        public void ObjectVariableNotSet_IgnoreQuickFixWorks()
        {
            var inputCode =
                @"
Private Sub Workbook_Open()
    
    Dim target As Range
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";
            var expectedCode =
                @"
Private Sub Workbook_Open()
    
    Dim target As Range
'@Ignore ObjectVariableNotSet
    target = Range(""A1"")
    
    target.Value = ""forgot something?""

End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObjectVariableNotSetInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCallStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                var fix = new IgnoreOnceQuickFix(state, new[] { inspection });
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteCommentSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteErrorSyntaxInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteGlobalInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteLetStatementInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new IgnoreOnceQuickFix(state, new[] { inspection });
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new OptionBaseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ParameterCanBeByValInspection(state);
                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ParameterNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ProcedureNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }


        [Test]
        [Category("QuickFixes")]
        public void ProcedureShouldBeFunction_IgnoreQuickFixWorks()
        {
            const string inputCode =
                @"Private Sub Foo(ByRef arg1 As Integer)
End Sub";

            const string expectedCode =
                @"'@Ignore ProcedureCanBeWrittenAsFunction
Private Sub Foo(ByRef arg1 As Integer)
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ProcedureCanBeWrittenAsFunctionInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new RedundantByRefModifierInspection(state) { Severity = CodeInspectionSeverity.Hint };
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var component = project.Object.VBComponents[0];
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new SelfAssignedDeclarationInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new UnassignedVariableUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        [Ignore("Todo")] // not sure how to handle GetBuiltInDeclarations
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
                .AddReference("VBA", MockVbeBuilder.LibraryPathVBA, 4, 1, true)
                .Build();
            var vbe = builder.AddProject(project).Build();

            var component = project.Object.VBComponents[0];
            var parser = MockParser.Create(vbe.Object);
            using (var state = parser.State)
            {
                // FIXME reinstate and unignore tests
                // refers to "UntypedFunctionUsageInspectionTests.GetBuiltInDeclarations()"
                //GetBuiltInDeclarations().ForEach(d => state.AddDeclaration(d));

                parser.Parse(new CancellationTokenSource());
                if (state.Status >= ParserState.Error)
                {
                    Assert.Inconclusive("Parser Error");
                }

                var inspection = new UntypedFunctionUsageInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
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

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new UseMeaningfulNameInspection(state, UseMeaningfulNameInspectionTests.GetInspectionSettings().Object);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());
                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotAssignedInspection(state);
                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableNotUsedInspection(state);
                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new VariableTypeNotDeclaredInspection(state);
                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspection.GetInspectionResults().First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var component = project.Object.VBComponents[0];
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new WriteOnlyPropertyInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var builder = new MockVbeBuilder();
            var project = builder.ProjectBuilder("VBAProject", ProjectProtection.Unprotected)
                .AddComponent("MyClass", ComponentType.ClassModule, inputCode)
                .Build();
            var component = project.Object.VBComponents[0];
            var vbe = builder.AddProject(project).Build();

            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var inspection = new BooleanAssignedInIfElseInspection(state);
                var inspector = InspectionsHelper.GetInspector(inspection);
                var inspectionResults = inspector.FindIssuesAsync(state, CancellationToken.None).Result;

                new IgnoreOnceQuickFix(state, new[] { inspection }).Fix(inspectionResults.First());

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}

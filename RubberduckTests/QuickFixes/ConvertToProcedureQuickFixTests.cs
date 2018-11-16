﻿using System.Linq;
using System.Threading;
using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using RubberduckTests.Mocks;

namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class ConvertToProcedureQuickFixTests : QuickFixTestBase
    {
        [Test]
        [Category("QuickFixes")]
        public void FunctionReturnValueNotUsed_QuickFixWorks_NoInterface()
        {
            const string inputCode =
                @"Public Function Foo(ByVal bar As String) As Boolean
    If True Then
        Foo = _
        True
    Else
        Foo = False
    End If
End Function

Public Sub Test()
    Foo ""test""
End Sub
";

            const string expectedCode =
                @"Public Sub Foo(ByVal bar As String)
    If True Then
        
    Else
        
    End If
End Sub

Public Sub Test()
    Foo ""test""
End Sub
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new FunctionReturnValueNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void FunctionReturnValueNotUsed_QuickFixWorks_NoInterface_ManyBodyStatements()
        {
            const string inputCode =
                @"Function foo(ByRef fizz As Boolean) As Boolean
    fizz = True
    goo
label1:
    foo = fizz
End Function

Sub goo()
End Sub

Public Sub Test()
    foo ""test""
End Sub
";

            const string expectedCode =
                @"Sub foo(ByRef fizz As Boolean)
    fizz = True
    goo
label1:
    
End Sub

Sub goo()
End Sub

Public Sub Test()
    foo ""test""
End Sub
";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new FunctionReturnValueNotUsedInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void FunctionReturnValueNotUsed_QuickFixWorks_Interface()
        {
            const string inputInterfaceCode =
                @"Public Function Test() As Integer
End Function";

            const string expectedInterfaceCode =
                @"Public Sub Test()
End Sub";

            const string inputImplementationCode1 =
                @"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string inputImplementationCode2 =
                @"Implements IFoo
Public Function IFoo_Test() As Integer
    IFoo_Test = 42
End Function";

            const string callSiteCode =
                @"
Public Function Baz()
    Dim testObj As IFoo
    Set testObj = new Bar
    testObj.Test
End Function";

            var builder = new MockVbeBuilder();
            var vbe = builder.ProjectBuilder("TestProject", ProjectProtection.Unprotected)
                .AddComponent("IFoo", ComponentType.ClassModule, inputInterfaceCode)
                .AddComponent("Bar", ComponentType.ClassModule, inputImplementationCode1)
                .AddComponent("Bar2", ComponentType.ClassModule, inputImplementationCode2)
                .AddComponent("TestModule", ComponentType.StandardModule, callSiteCode)
                .AddProjectToVbeBuilder().Build();

            var (state, rewritingManager) = MockParser.CreateAndParseWithRewritingManager(vbe.Object);
            using (state)
            {

                var inspection = new FunctionReturnValueNotUsedInspection(state);
                var inspectionResults = inspection.GetInspectionResults(CancellationToken.None);
                var rewriteSession = rewritingManager.CheckOutCodePaneSession();

                new ConvertToProcedureQuickFix().Fix(inspectionResults.First(), rewriteSession);

                var component = vbe.Object.VBProjects[0].VBComponents[0];
                var actualCode = rewriteSession.CheckOutModuleRewriter(component.QualifiedModuleName).GetText();
                Assert.AreEqual(expectedInterfaceCode, actualCode);
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void NonReturningFunction_QuickFixWorks_Function()
        {
            const string inputCode =
                @"Function Foo() As Boolean
End Function";

            const string expectedCode =
                @"Sub Foo()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new NonReturningFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void GivenFunctionNameWithTypeHint_SubNameHasNoTypeHint()
        {
            const string inputCode =
                @"Function Foo$()
End Function";

            const string expectedCode =
                @"Sub Foo()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new NonReturningFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonReturningFunction_QuickFixWorks_FunctionReturnsImplicitVariant()
        {
            const string inputCode =
                @"Function Foo()
End Function";

            const string expectedCode =
                @"Sub Foo()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new NonReturningFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void NonReturningFunction_QuickFixWorks_FunctionHasVariable()
        {
            const string inputCode =
                @"Function Foo(ByVal b As Boolean) As String
End Function";

            const string expectedCode =
                @"Sub Foo(ByVal b As Boolean)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new NonReturningFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void GivenNonReturningPropertyGetter_QuickFixConvertsToSub()
        {
            const string inputCode =
                @"Property Get Foo() As Boolean
End Property";

            const string expectedCode =
                @"Sub Foo()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new NonReturningFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void GivenNonReturningPropertyGetWithTypeHint_QuickFixDropsTypeHint()
        {
            const string inputCode =
                @"Property Get Foo$()
End Property";

            const string expectedCode =
                @"Sub Foo()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new NonReturningFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void GivenImplicitVariantPropertyGetter_StillConvertsToSub()
        {
            const string inputCode =
                @"Property Get Foo()
End Property";

            const string expectedCode =
                @"Sub Foo()
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new NonReturningFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void GivenParameterizedPropertyGetter_QuickFixKeepsParameter()
        {
            const string inputCode =
                @"Property Get Foo(ByVal b As Boolean) As String
End Property";

            const string expectedCode =
                @"Sub Foo(ByVal b As Boolean)
End Sub";

            var actualCode = ApplyQuickFixToFirstInspectionResult(inputCode, state => new NonReturningFunctionInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new ConvertToProcedureQuickFix();
        }
    }
}

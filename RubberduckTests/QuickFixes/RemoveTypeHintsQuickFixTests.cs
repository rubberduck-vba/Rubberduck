using NUnit.Framework;
using Rubberduck.CodeAnalysis.Inspections.Concrete;
using Rubberduck.CodeAnalysis.QuickFixes;
using Rubberduck.CodeAnalysis.QuickFixes.Concrete;
using Rubberduck.Parsing.VBA;
using RubberduckTests.Mocks;


namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveTypeHintsQuickFixTests : QuickFixTestBase
    {

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_LongTypeHint()
        {
            const string inputCode =
                @"Public Foo&";

            const string expectedCode =
                @"Public Foo As Long";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteTypeHintInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_IntegerTypeHint()
        {
            const string inputCode =
                @"Public Foo%";

            const string expectedCode =
                @"Public Foo As Integer";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteTypeHintInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_DoubleTypeHint()
        {
            const string inputCode =
                @"Public Foo#";

            const string expectedCode =
                @"Public Foo As Double";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteTypeHintInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_SingleTypeHint()
        {
            const string inputCode =
                @"Public Foo!";

            const string expectedCode =
                @"Public Foo As Single";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteTypeHintInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_DecimalTypeHint()
        {
            const string inputCode =
                @"Public Foo@";

            const string expectedCode =
                @"Public Foo As Currency";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteTypeHintInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_StringTypeHint()
        {
            const string inputCode =
                @"Public Foo$";

            const string expectedCode =
                @"Public Foo As String";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteTypeHintInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Function_StringTypeHint()
        {
            const string inputCode =
                @"Public Function Foo$(ByVal fizz As Integer)
    Foo = ""test""
End Function";

            const string expectedCode =
                @"Public Function Foo(ByVal fizz As Integer) As String
    Foo = ""test""
End Function";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteTypeHintInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_PropertyGet_StringTypeHint()
        {
            const string inputCode =
                @"Public Property Get Foo$(ByVal fizz As Integer)
    Foo = ""test""
End Property";

            const string expectedCode =
                @"Public Property Get Foo(ByVal fizz As Integer) As String
    Foo = ""test""
End Property";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteTypeHintInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Parameter_StringTypeHint()
        {
            const string inputCode =
                @"Public Sub Foo(ByVal fizz$)
    Foo = ""test""
End Sub";

            const string expectedCode =
                @"Public Sub Foo(ByVal fizz As String)
    Foo = ""test""
End Sub";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteTypeHintInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Variable_StringTypeHint()
        {
            const string inputCode =
                @"Public Sub Foo()
    Dim buzz$
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Dim buzz As String
End Sub";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteTypeHintInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Constant_StringTypeHint()
        {
            const string inputCode =
                @"Public Sub Foo()
    Const buzz$ = """"
End Sub";

            const string expectedCode =
                @"Public Sub Foo()
    Const buzz As String = """"
End Sub";

            var actualCode = ApplyQuickFixToAllInspectionResults(inputCode, state => new ObsoleteTypeHintInspection(state));
            Assert.AreEqual(expectedCode, actualCode);
        }


        protected override IQuickFix QuickFix(RubberduckParserState state)
        {
            return new RemoveTypeHintsQuickFix();
        }
    }
}

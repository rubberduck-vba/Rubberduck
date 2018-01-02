using NUnit.Framework;
using Rubberduck.Inspections.Concrete;
using Rubberduck.Inspections.QuickFixes;
using RubberduckTests.Mocks;


namespace RubberduckTests.QuickFixes
{
    [TestFixture]
    public class RemoveTypeHintsQuickFixTests
    {

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_LongTypeHint()
        {
            const string inputCode =
                @"Public Foo&";

            const string expectedCode =
                @"Public Foo As Long";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new RemoveTypeHintsQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_IntegerTypeHint()
        {
            const string inputCode =
                @"Public Foo%";

            const string expectedCode =
                @"Public Foo As Integer";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new RemoveTypeHintsQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_DoubleTypeHint()
        {
            const string inputCode =
                @"Public Foo#";

            const string expectedCode =
                @"Public Foo As Double";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new RemoveTypeHintsQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_SingleTypeHint()
        {
            const string inputCode =
                @"Public Foo!";

            const string expectedCode =
                @"Public Foo As Single";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new RemoveTypeHintsQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_DecimalTypeHint()
        {
            const string inputCode =
                @"Public Foo@";

            const string expectedCode =
                @"Public Foo As Decimal";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new RemoveTypeHintsQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

        [Test]
        [Category("QuickFixes")]
        public void ObsoleteTypeHint_QuickFixWorks_Field_StringTypeHint()
        {
            const string inputCode =
                @"Public Foo$";

            const string expectedCode =
                @"Public Foo As String";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new RemoveTypeHintsQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new RemoveTypeHintsQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new RemoveTypeHintsQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new RemoveTypeHintsQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new RemoveTypeHintsQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
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

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {

                var inspection = new ObsoleteTypeHintInspection(state);
                var inspectionResults = inspection.GetInspectionResults();

                var fix = new RemoveTypeHintsQuickFix(state);
                foreach (var result in inspectionResults)
                {
                    fix.Fix(result);
                }

                Assert.AreEqual(expectedCode, state.GetRewriter(component).GetText());
            }
        }

    }
}

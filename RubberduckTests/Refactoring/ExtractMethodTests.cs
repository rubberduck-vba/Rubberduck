using System;
using Moq;
using NUnit.Framework;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ExtractMethodTests
    {
        private string Indent(IVBE vbe, string code)
        {
            var indenter = MockIndenterBuilder.Create(vbe);
            return Indent(indenter, code);
        }

        private string Indent(IIndenter indenter, string code)
        {
            return string.Join(Environment.NewLine,
                indenter.Indent(code.Split(new[] {Environment.NewLine}, StringSplitOptions.None)));
        }

        private ExtractMethodModel BuildModel(RubberduckParserState state, IVBComponent component, Selection selection)
        {
            var qualifiedSelection = MockQualifiedSelectionBuilder.CreateQualifiedSelection(component, selection);

            var validator = new ExtractMethodSelectionValidation(state.AllDeclarations, component.CodeModule);
            validator.ValidateSelection(qualifiedSelection);
            var indenter = MockIndenterBuilder.Create(component.VBE);
            return new ExtractMethodModel(state, qualifiedSelection, validator.SelectedContexts, indenter, component.CodeModule);
        }

        [Test]
        [Category("ExtractMethodModel_PreviewCode")]
        public void ExtractMethodModel_PreviewCode_Base()
        {
            var inputCode = 
@"Option Explicit

Public Sub Test()
  Dim SomeString As String
  Dim OtherString As String

  SomeString = ""Hello, world!""
  OtherString = ""Goodbye, world!""
  
  Msgbox SomeString
  Msgbox OtherString
End Sub";
            var selection = new Selection(11, 1, 11, 21);

            var expectedPreviewCode = 
@"Private Sub NewMethod(OtherString As String)
  Msgbox OtherString
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection, true);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var model = BuildModel(state, component, selection);
                Assert.AreEqual(Indent(vbe.Object, expectedPreviewCode), model.PreviewCode);
            }
        }

        [Test]
        [Category("ExtractMethodModel_PreviewCode")]
        public void ExtractMethodModel_PreviewCode_ExcludeDim()
        {
            var inputCode =
@"Option Explicit

Public Sub Test()
  Dim SomeString As String
  SomeString = ""Hello, world!""
  Msgbox SomeString


  Dim OtherString As String  
  OtherString = ""Goodbye, world!""
  Msgbox OtherString
End Sub";
            var selection = new Selection(4, 1, 6, 20);

            var expectedPreviewCode =
@"Private Sub NewMethod(SomeString As String)
  Msgbox SomeString
End Sub";

            var vbe = MockVbeBuilder.BuildFromSingleStandardModule(inputCode, out var component, selection, true);
            using (var state = MockParser.CreateAndParse(vbe.Object))
            {
                var model = BuildModel(state, component, selection);
                Assert.AreEqual(Indent(vbe.Object, expectedPreviewCode), model.PreviewCode);
            }
        }
    }
}

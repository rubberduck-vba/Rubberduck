using System;
using System.Linq;
using NUnit.Framework;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.ExtractMethod;
using Rubberduck.SmartIndenter;
using Rubberduck.VBEditor;
using RubberduckTests.Mocks;

namespace RubberduckTests.Refactoring
{
    [TestFixture]
    public class ExtractMethodTests
    {
        private string Indent(IIndenter indenter, string code)
        {
            return string.Join(Environment.NewLine,
                indenter.Indent(code.Split(new[] {Environment.NewLine}, StringSplitOptions.None)));
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
                var qualifiedSelection = MockQualifiedSelectionBuilder.CreateQualifiedSelection(component, selection);
                
                var validator = new ExtractMethodSelectionValidation(state.AllDeclarations, component.CodeModule);
                validator.ValidateSelection(qualifiedSelection);
                var indenter = MockIndenterBuilder.Create(vbe.Object);
                var model = new ExtractMethodModel(state, qualifiedSelection, validator.SelectedContexts, indenter, component.CodeModule);

                Assert.AreEqual(model.PreviewCode, Indent(indenter, expectedPreviewCode));
            }
        }
    }
}

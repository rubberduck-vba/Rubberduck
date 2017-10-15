using System.Linq;
using System.Runtime.InteropServices;
using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UnitTesting;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    /// <summary>
    /// A command that adds a new test method stub to the active code pane.
    /// </summary>
    [ComVisible(false)]
    public class AddTestMethodExpectedErrorCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly RubberduckParserState _state;

        public AddTestMethodExpectedErrorCommand(IVBE vbe, RubberduckParserState state) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
        }

        public const string NamePlaceholder = "%METHODNAME%";
        private const string TestMethodBaseName = "TestMethod";

        public static readonly string TestMethodExpectedErrorTemplate = string.Concat(
            "'@TestMethod\r\n",
            "Public Sub ", NamePlaceholder, "() 'TODO ", RubberduckUI.UnitTest_NewMethod_Rename, "\r\n",
            "    Const ExpectedError As Long = 0 'TODO ", RubberduckUI.UnitTest_NewMethod_ChangeErrorNo, "\r\n",
            "    On Error GoTo TestFail\r\n",
            "    \r\n",
            "    'Arrange:\r\n\r\n",
            "    'Act:\r\n\r\n",
            "Assert:\r\n",
            "    Assert.Fail \"", RubberduckUI.UnitTest_NewMethod_ErrorNotRaised, ".\"\r\n\r\n",
            "TestExit:\r\n",
            "    Exit Sub\r\n",
            "TestFail:\r\n",
            "    If Err.Number = ExpectedError Then\r\n",
            "        Resume TestExit\r\n",
            "    Else\r\n",
            "        Resume Assert\r\n",
            "    End If\r\n",
            "End Sub\r\n"
            );

        protected override bool CanExecuteImpl(object parameter)
        {
            var pane = _vbe.ActiveCodePane;
            {
                if (_state.Status != ParserState.Ready || pane.IsWrappingNullReference)
                {
                    return false;
                }

                var testModules = _state.AllUserDeclarations.Where(d =>
                            d.DeclarationType == DeclarationType.ProceduralModule &&
                            d.Annotations.Any(a => a.AnnotationType == AnnotationType.TestModule));

                try
                {
                    // the code modules consistently match correctly, but the components don't
                    var component = _vbe.SelectedVBComponent;
                    var selectedModule = component.CodeModule;
                    {
                        var result = testModules.Any( a => a.QualifiedName.QualifiedModuleName.Component.CodeModule.Equals(selectedModule));
                        return result;
                    }
                }
                catch (COMException)
                {
                    return false;
                }
            }
        }

        protected override void ExecuteImpl(object parameter)
        {
            var pane = _vbe.ActiveCodePane;
            if (pane.IsWrappingNullReference)
            {
                return;
            }

            var activeModule = pane.CodeModule;
            var declaration = _state.GetTestModules().FirstOrDefault(f =>
            {
                var thisModule = f.QualifiedName.QualifiedModuleName.Component.CodeModule;
                return thisModule.Equals(activeModule);
            });

            if (declaration != null)
            {
                var module = pane.CodeModule;
                var name = GetNextTestMethodName(module.Parent);
                var body = TestMethodExpectedErrorTemplate.Replace(NamePlaceholder, name);
                module.InsertLines(module.CountOfLines, body);
            }

            _state.OnParseRequested(this, _vbe.SelectedVBComponent);
        }

        private string GetNextTestMethodName(IVBComponent component)
        {
            var names = component.GetTests(_vbe, _state).Select(test => test.Declaration.IdentifierName);
            var index = names.Count(n => n.StartsWith(TestMethodBaseName)) + 1;

            return string.Concat(TestMethodBaseName, index);
        }
    }
}

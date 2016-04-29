using System;
using System.Diagnostics;
using Microsoft.Office.Core;
using Microsoft.Vbe.Interop;

namespace Rubberduck.VBEditor.VBEHost
{
    public sealed class FallbackApp : IHostApplication
    {
        private readonly VBE _vbe;
        private readonly CommandBarButton _runButton;

        private const int DebugCommandBarId = 4;
        private const int RunMacroCommand = 186;

        public FallbackApp(VBE vbe)
        {
            _vbe = vbe;
            var mainCommandBar = _vbe.CommandBars[DebugCommandBarId];
            _runButton = (CommandBarButton)mainCommandBar.FindControl(Id: RunMacroCommand);
        }

        public void Run(QualifiedMemberName qualifiedMemberName)
        {
            var component = qualifiedMemberName.QualifiedModuleName.Component;
            var line = component.CodeModule.get_ProcBodyLine(qualifiedMemberName.MemberName, vbext_ProcKind.vbext_pk_Proc);

            component.CodeModule.CodePane.SetSelection(line, 1, line, 1);
            component.CodeModule.CodePane.ForceFocus();

            _runButton.Execute();
            // note: this can't work... because the .Execute() call isn't blocking, so method returns before test method actually runs.
        }

        public TimeSpan TimedMethodCall(QualifiedMemberName qualifiedMemberName)
        {
            var stopwatch = Stopwatch.StartNew();

            Run(qualifiedMemberName);

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public string ApplicationName { get { return "(unknown)"; } }

        public void Dispose()
        {
        }
    }
}
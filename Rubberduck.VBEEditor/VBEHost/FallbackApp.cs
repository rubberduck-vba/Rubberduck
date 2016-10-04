using System;
using System.Diagnostics;
using Microsoft.Office.Core;
using Rubberduck.VBEditor.DisposableWrappers;
using Rubberduck.VBEditor.DisposableWrappers.VBA;
using Rubberduck.VBEditor.Extensions;

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
            using (var module = component.CodeModule)
            {
                var line = module.GetProcBodyStartLine(qualifiedMemberName.MemberName, ProcKind.Procedure);
                using (var pane = module.CodePane)
                {
                    pane.SetSelection(line, 1, line, 1);
                    pane.ForceFocus();
                }
            }

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

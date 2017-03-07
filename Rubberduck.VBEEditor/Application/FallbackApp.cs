using System;
using System.Diagnostics;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core.Abstract;

namespace Rubberduck.VBEditor.Application
{
    public sealed class FallbackApp : IHostApplication
    {
        private readonly ICommandBarButton _runButton;

        private const int DebugCommandBarId = 4;
        private const int RunMacroCommand = 186;

        public FallbackApp(IVBE vbe)
        {
            var mainCommandBar = vbe.CommandBars[DebugCommandBarId];
            _runButton = (ICommandBarButton)mainCommandBar.FindControl(RunMacroCommand);
        }

        public void Run(dynamic declaration)
        {
            var qualifiedMemberName = declaration.QualifiedName;
            var component = qualifiedMemberName.QualifiedModuleName.Component;
            var module = component.CodeModule;
            {
                var line = module.GetProcBodyStartLine(qualifiedMemberName.MemberName, ProcKind.Procedure);
                var pane = module.CodePane;
                {
                    pane.Selection = new Selection(line, 1, line, 1);
                }
            }

            _runButton.Execute();
            // note: this can't work... because the .Execute() call isn't blocking, so method returns before test method actually runs.
        }

        public object Run(string name, object[] args)
        {
            return null;
        }

        public TimeSpan TimedMethodCall(dynamic declaration)
        {
            var stopwatch = Stopwatch.StartNew();

            Run(declaration);

            stopwatch.Stop();
            return stopwatch.Elapsed;
        }

        public string ApplicationName { get { return "(unknown)"; } }

        public void Dispose()
        {
        }
    }
}

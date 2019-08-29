using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using Rubberduck.Interaction;
using Rubberduck.Parsing.VBA;
using Rubberduck.Resources;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.VBEditor.ComManagement.TypeLibs.Abstract;
using Rubberduck.VBEditor.Events;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.VbeRuntime.Settings;

namespace Rubberduck.UI.Command.ComCommands
{
    [ComVisible(false)]
    public class ReparseCancellationFlag
    {
        public bool Canceled { get; set; }
    }

    [ComVisible(false)]
    public class ReparseCommand : ComCommandBase
    {
        private readonly IVBE _vbe;
        private readonly IVBETypeLibsAPI _typeLibApi;
        private readonly IVbeSettings _vbeSettings;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private readonly GeneralSettings _settings;

        public ReparseCommand(
            IVBE vbe, 
            IConfigurationService<GeneralSettings> settingsProvider, 
            RubberduckParserState state, 
            IVBETypeLibsAPI typeLibApi, 
            IVbeSettings vbeSettings, 
            IMessageBox messageBox, 
            IVbeEvents vbeEvents) 
            : base(vbeEvents)
        {
            _vbe = vbe;
            _vbeSettings = vbeSettings;
            _typeLibApi = typeLibApi;
            _state = state;
            _settings = settingsProvider.Read();
            _messageBox = messageBox;

            AddToCanExecuteEvaluation(SpecialEvaluateCanExecute);
        }

        private bool SpecialEvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Pending
                   || _state.Status == ParserState.Ready
                   || _state.Status == ParserState.Error
                   || _state.Status == ParserState.ResolverError
                   || _state.Status == ParserState.UnexpectedError;
        }

        protected override void OnExecute(object parameter)
        {
            // WPF binds to EvaluateCanExecute asychronously, which means that in some instances the bound refresh control will
            // enable itself based on a "stale" ParserState. There's no easy way to test for race conditions inside WPF, so we
            // need to make this test again...
            if (!CanExecute(parameter))
            {
                return;
            }

            if (_settings.CompileBeforeParse)
            {
                if (!VerifyCompileOnDemand())
                {
                    if (parameter is ReparseCancellationFlag cancellation)
                    {
                        cancellation.Canceled = true;
                    }
                    return;
                }

                if (AreAllProjectsCompiled(out var failedNames))
                {
                    if (!PromptUserToContinue(failedNames))
                    {
                        if (parameter is ReparseCancellationFlag cancellation)
                        {
                            cancellation.Canceled = true;
                        }
                        return;
                    }
                }
            }
            _state.OnParseRequested(this);
        }

        private bool VerifyCompileOnDemand()
        {
            if (_vbeSettings.CompileOnDemand)
            {
                return _messageBox.ConfirmYesNo(RubberduckUI.Command_Reparse_CompileOnDemandEnabled, RubberduckUI.Command_Reparse_CompileOnDemandEnabled_Caption, false);
            }

            return true;
        }

        private bool AreAllProjectsCompiled(out List<string> failedNames)
        {
            failedNames = new List<string>();
            using (var projects = _vbe.VBProjects)
            {
                foreach (var project in projects)
                {
                    using (project)
                    {
                        if (project.Protection != ProjectProtection.Unprotected)
                        {
                            continue;
                        }

                        if (!_typeLibApi.CompileProject(project))
                        {
                            failedNames.Add(project.Name);
                        }
                    }
                }
            }

            return failedNames.Any();
        }

        private bool PromptUserToContinue(List<string> failedNames)
        {
            var formattedList = string.Concat(Environment.NewLine, Environment.NewLine,
                string.Join(Environment.NewLine, failedNames));
            // FIXME using Exclamation instead of warning now... 
            return _messageBox.ConfirmYesNo(
                string.Format(RubberduckUI.Command_Reparse_CannotCompile,
                    formattedList),
                RubberduckUI.Command_Reparse_CannotCompile_Caption, false);
            
        }
    }
}

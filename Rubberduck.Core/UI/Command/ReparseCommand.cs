using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.Interaction;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.Resources;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.VbeRuntime.Settings;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    public class ReparseCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly IVBETypeLibsAPI _typeLibApi;
        private readonly IVbeSettings _vbeSettings;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private readonly GeneralSettings _settings;

        public ReparseCommand(IVBE vbe, IConfigProvider<GeneralSettings> settingsProvider, RubberduckParserState state, IVBETypeLibsAPI typeLibApi, IVbeSettings vbeSettings, IMessageBox messageBox) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _vbeSettings = vbeSettings;
            _typeLibApi = typeLibApi;
            _state = state;
            _settings = settingsProvider.Create();
            _messageBox = messageBox;
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Pending
                   || _state.Status == ParserState.Ready
                   || _state.Status == ParserState.Error
                   || _state.Status == ParserState.ResolverError
                   || _state.Status == ParserState.UnexpectedError;
        }

        protected override void OnExecute(object parameter)
        {
            if (_settings.CompileBeforeParse)
            {
                if (!VerifyCompileOnDemand())
                {
                    return;
                }

                if (AreAllProjectsCompiled(out var failedNames))
                {
                    if (!PromptUserToContinue(failedNames))
                    {
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

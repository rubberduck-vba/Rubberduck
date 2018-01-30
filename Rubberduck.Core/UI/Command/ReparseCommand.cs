using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using NLog;
using Rubberduck.Parsing.VBA;
using Rubberduck.Settings;
using Rubberduck.SettingsProvider;
using Rubberduck.UI.CodeExplorer.Commands;
using Rubberduck.VBEditor.ComManagement.TypeLibsAPI;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command
{
    [ComVisible(false)]
    [CodeExplorerCommand]
    public class ReparseCommand : CommandBase
    {
        private readonly IVBE _vbe;
        private readonly IVBETypeLibsAPI _typeLibApi;
        private readonly IMessageBox _messageBox;
        private readonly RubberduckParserState _state;
        private readonly GeneralSettings _settings;

        public ReparseCommand(IVBE vbe, IConfigProvider<GeneralSettings> settingsProvider, RubberduckParserState state, IVBETypeLibsAPI typeLibApi, IMessageBox messageBox) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
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
                if (CompileAllProjects(out var failedNames))
                {
                    if (!PromptUserToContinue(failedNames))
                    {
                        return;
                    }
                }
            }
            _state.OnParseRequested(this);
        }

        private bool CompileAllProjects(out List<string> failedNames)
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
            var result = _messageBox.Show(
                string.Format(RubberduckUI.Command_Reparse_CannotCompile,
                    formattedList),
                RubberduckUI.Command_Reparse_CannotCompile_Caption, MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            return result == DialogResult.Yes;
        }
    }
}

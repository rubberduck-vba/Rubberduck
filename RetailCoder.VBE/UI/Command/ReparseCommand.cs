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
        private readonly RubberduckParserState _state;
        private readonly GeneralSettings _settings;

        public ReparseCommand(IVBE vbe, IConfigProvider<GeneralSettings> settingsProvider, RubberduckParserState state) : base(LogManager.GetCurrentClassLogger())
        {
            _vbe = vbe;
            _state = state;
            _settings = settingsProvider.Create();
        }

        protected override bool EvaluateCanExecute(object parameter)
        {
            return _state.Status == ParserState.Pending
                   || _state.Status == ParserState.Ready
                   || _state.Status == ParserState.Error
                   || _state.Status == ParserState.ResolverError;
        }

        protected override void OnExecute(object parameter)
        {
            if (_settings.CompileBeforeParse)
            {
                var failedNames = new List<string>();
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

                            if (!VBETypeLibsAPI.CompileProject(project))
                            {
                                failedNames.Add(project.Name);
                            }
                        }
                    }
                }

                if (failedNames.Any())
                {
                    var formattedList = string.Concat(Environment.NewLine, Environment.NewLine,
                        string.Join(Environment.NewLine, failedNames));
                    var msgbox = new MessageBox();
                    var result = msgbox.Show(
                        string.Format(RubberduckUI.Command_Reparse_CannotCompile,
                            formattedList),
                        RubberduckUI.Command_Reparse_CannotCompile_Caption, MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    if (result != DialogResult.Yes)
                    {
                        return;
                    }
                }
            }
            _state.OnParseRequested(this);
        }
    }
}

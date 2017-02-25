using System;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    public interface IFolderBrowser : IDisposable
    {
        string Description { get; set; }
        bool ShowNewFolderButton { get; set; }
        string RootFolder { get; set; }
        string SelectedPath { get; set; }
        DialogResult ShowDialog();
    }

    public class FolderBrowser : IFolderBrowser
    {
        private readonly FolderBrowserDialog _dialog;
        private readonly IEnvironmentProvider _environment;

        public FolderBrowser(IEnvironmentProvider environment, string description, bool showNewFolderButton, string rootFolder)
        {
            _environment = environment;
            _dialog = new FolderBrowserDialog
            {
                Description = description,
                SelectedPath = rootFolder,
                ShowNewFolderButton = showNewFolderButton
            };
        }

        public FolderBrowser(IEnvironmentProvider environment, string description, bool showNewFolderButton)
            : this(environment, description, showNewFolderButton, environment.GetFolderPath(Environment.SpecialFolder.MyDocuments))
        { }

        public FolderBrowser(IEnvironmentProvider environment, string description)
            : this(environment, description, true)
        { }

        public string Description
        {
            get { return _dialog.Description; }
            set { _dialog.Description = value; }
        }

        public bool ShowNewFolderButton
        {
            get { return _dialog.ShowNewFolderButton; }
            set { _dialog.ShowNewFolderButton = value; }
        }

        public string RootFolder
        {
            get { return _dialog.SelectedPath; }
            set { _dialog.SelectedPath = value; }
        }

        public string SelectedPath
        {
            get { return _dialog.SelectedPath; }
            set { _dialog.SelectedPath = value; }
        }

        public DialogResult ShowDialog()
        {
            return _dialog.ShowDialog();
        }

        public void Dispose()
        {
            if (_dialog != null)
            {
                _dialog.Dispose();
            }
        }
    }
}

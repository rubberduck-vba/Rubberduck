using System;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    public interface IFolderBrowser : IDisposable
    {
        string Description { get; set; }
        bool ShowNewFolderButton { get; set; }
        Environment.SpecialFolder RootFolder { get; set; }
        string SelectedPath { get; set; }
        DialogResult ShowDialog();
    }

    public class FolderBrowser : IFolderBrowser
    {
        private readonly FolderBrowserDialog _dialog;

        public FolderBrowser(string description, bool showNewFolderButton, Environment.SpecialFolder rootFolder)
        {
            _dialog = new FolderBrowserDialog
            {
                Description = description,
                RootFolder = rootFolder,
                ShowNewFolderButton = showNewFolderButton
            };
        }

        public FolderBrowser(string description, bool showNewFolderButton)
            :this(description, showNewFolderButton, Environment.SpecialFolder.MyDocuments)
        { }

        public FolderBrowser(string description)
            : this(description, true)
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

        public Environment.SpecialFolder RootFolder
        {
            get { return _dialog.RootFolder; }
            set { _dialog.RootFolder = value; }
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

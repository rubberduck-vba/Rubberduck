using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Forms = System.Windows.Forms;
using System.ComponentModel;

namespace Rubberduck.UI
{
    public interface IFolderBrowserDialog : IDisposable
    {
        string Description { get; set; }
        Environment.SpecialFolder RootFolder { get; set; }
        string SelectedPath { get; set; }
        bool ShowNewFolderButton { get; set; }
        object Tag { get; set; }

        void Reset();
        Forms.DialogResult ShowDialog();
    }

    /// <summary>
    /// Mockable FolderBrowswerDialog that simply wraps the System.Windows.Forms.FolderDialogBrowswer
    /// </summary>
    public sealed class FolderBrowserDialog : IFolderBrowserDialog
    {
        private Forms.FolderBrowserDialog _dialog;

        public FolderBrowserDialog()
        {
            _dialog = new Forms.FolderBrowserDialog();
        }

        public string Description
        {
            get { return _dialog.Description;}
            set { _dialog.Description = value; }
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

        public bool ShowNewFolderButton
        {
            get { return _dialog.ShowNewFolderButton; }
            set { _dialog.ShowNewFolderButton = value; }
        }

        public object Tag
        {
            get { return _dialog.Tag; }
            set { _dialog.Tag = value; }
        }

        public void Reset()
        {
            _dialog.Reset();
        }

        public Forms.DialogResult ShowDialog()
        {
            return _dialog.ShowDialog();
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        void Dispose(bool disposing)
        {
            if (disposing)
            {
                _dialog.Dispose();
                _dialog = null;
            }
        }
    }
}

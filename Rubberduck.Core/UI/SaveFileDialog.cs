using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.Remoting;
using System.Windows.Forms;

namespace Rubberduck.UI
{
    public interface ISaveFileDialog : IDisposable
    {
        event EventHandler Disposed;
        event EventHandler FileOk;
        event EventHandler HelpRequest;

        bool AddExtension { get; set; }
        bool AutoUpgradeEnabled { get; set; }
        bool CheckFileExists { get; set; }
        bool CheckPathExists { get; set; }
        bool CreatePrompt { get; set; }
        string DefaultExt { get; set; }
        bool DereferenceLinks { get; set; }
        string FileName { get; set; }
        string Filter { get; set; }
        int FilterIndex { get; set; }
        string InitialDirectory { get; set; }
        bool OverwritePrompt { get; set; }
        bool RestoreDirectory { get; set; }
        bool ShowHelp { get; set; }
        ISite Site { get; set; }
        bool SupportMultiDottedExtensions { get; set; }
        object Tag { get; set; }
        string Title { get; set; }
        bool ValidateNames { get; set; }

        IContainer Container { get; }
        FileDialogCustomPlacesCollection CustomPlaces { get; }
        string[] FileNames { get; }

        ObjRef CreateObjRef(Type requestedType);
        object GetLifetimeService();
        object InitializeLifetimeService();
        Stream OpenFile();
        void Reset();
        DialogResult ShowDialog();
    }

    public class SaveFileDialog : ISaveFileDialog
    {
        private readonly System.Windows.Forms.SaveFileDialog _saveFileDialog;

        public SaveFileDialog()
        {
            _saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            _saveFileDialog.Disposed += SaveFileDialog_Disposed;
            _saveFileDialog.FileOk += SaveFileDialog_FileOk;
            _saveFileDialog.HelpRequest += SaveFileDialog_HelpRequest;
        }

        public bool AddExtension
        {
            get => _saveFileDialog.AddExtension;
            set => _saveFileDialog.AddExtension = value;
        }

        public bool AutoUpgradeEnabled
        {
            get => _saveFileDialog.AutoUpgradeEnabled;
            set => _saveFileDialog.AutoUpgradeEnabled = value;
        }

        public bool CheckFileExists
        {
            get => _saveFileDialog.CheckFileExists;
            set => _saveFileDialog.CheckFileExists = value;
        }

        public bool CheckPathExists
        {
            get => _saveFileDialog.CheckPathExists;
            set => _saveFileDialog.CheckPathExists = value;
        }

        public IContainer Container => _saveFileDialog.Container;

        public ObjRef CreateObjRef(Type requestedType)
        {
            return _saveFileDialog.CreateObjRef(requestedType);
        }

        public bool CreatePrompt
        {
            get => _saveFileDialog.CreatePrompt;
            set => _saveFileDialog.CreatePrompt = value;
        }

        public FileDialogCustomPlacesCollection CustomPlaces => _saveFileDialog.CustomPlaces;

        public string DefaultExt
        {
            get => _saveFileDialog.DefaultExt;
            set => _saveFileDialog.DefaultExt = value;
        }

        public bool DereferenceLinks
        {
            get => _saveFileDialog.DereferenceLinks;
            set => _saveFileDialog.DereferenceLinks = value;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        private bool _isDisposed;
        protected virtual void Dispose(bool disposing)
        {
            if (_isDisposed || !disposing)
            {
                return;
            }

            _saveFileDialog.Dispose();
            _isDisposed = true;
        }

        public event EventHandler Disposed;
        private void SaveFileDialog_Disposed(object sender, EventArgs e)
        {
            Disposed?.Invoke(sender, e);
        }

        public override bool Equals(object obj)
        {
            return _saveFileDialog.Equals(obj);
        }

        public string FileName
        {
            get => _saveFileDialog.FileName;
            set => _saveFileDialog.FileName = value;
        }

        public string[] FileNames => _saveFileDialog.FileNames;

        public event EventHandler FileOk;
        private void SaveFileDialog_FileOk(object sender, EventArgs e)
        {
            FileOk?.Invoke(sender, e);
        }

        public string Filter
        {
            get => _saveFileDialog.Filter;
            set => _saveFileDialog.Filter = value;
        }

        public int FilterIndex
        {
            get => _saveFileDialog.FilterIndex;
            set => _saveFileDialog.FilterIndex = value;
        }

        public override int GetHashCode()
        {
            return _saveFileDialog.GetHashCode();
        }

        public object GetLifetimeService()
        {
            return _saveFileDialog.GetLifetimeService();
        }

        public event EventHandler HelpRequest;
        private void SaveFileDialog_HelpRequest(object sender, EventArgs e)
        {
            HelpRequest?.Invoke(sender, e);
        }

        public string InitialDirectory
        {
            get => _saveFileDialog.InitialDirectory;
            set => _saveFileDialog.InitialDirectory = value;
        }

        public object InitializeLifetimeService()
        {
            return _saveFileDialog.InitializeLifetimeService();
        }

        public Stream OpenFile()
        {
            return _saveFileDialog.OpenFile();
        }

        public bool OverwritePrompt
        {
            get => _saveFileDialog.OverwritePrompt;
            set => _saveFileDialog.OverwritePrompt = value;
        }

        public void Reset()
        {
            _saveFileDialog.Reset();
        }

        public bool RestoreDirectory
        {
            get => _saveFileDialog.RestoreDirectory;
            set => _saveFileDialog.RestoreDirectory = value;
        }

        public DialogResult ShowDialog()
        {
            return _saveFileDialog.ShowDialog();
        }

        public bool ShowHelp
        {
            get => _saveFileDialog.ShowHelp;
            set => _saveFileDialog.ShowHelp = value;
        }

        public ISite Site
        {
            get => _saveFileDialog.Site;
            set => _saveFileDialog.Site = value;
        }

        public bool SupportMultiDottedExtensions
        {
            get => _saveFileDialog.SupportMultiDottedExtensions;
            set => _saveFileDialog.SupportMultiDottedExtensions = value;
        }

        public object Tag
        {
            get => _saveFileDialog.Tag;
            set => _saveFileDialog.Tag = value;
        }

        public string Title
        {
            get => _saveFileDialog.Title;
            set => _saveFileDialog.Title = value;
        }

        public override string ToString()
        {
            return _saveFileDialog.ToString();
        }

        public bool ValidateNames
        {
            get => _saveFileDialog.ValidateNames;
            set => _saveFileDialog.ValidateNames = value;
        }
    }
}

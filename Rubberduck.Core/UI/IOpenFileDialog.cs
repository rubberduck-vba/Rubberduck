using System;
using System.ComponentModel;
using System.IO;
using System.Runtime.Remoting;
using System.Windows.Forms;
// ReSharper disable EventNeverSubscribedTo.Global
// ReSharper disable UnusedMember.Global

namespace Rubberduck.UI
{
    public interface IOpenFileDialog : IDisposable
    {
        event EventHandler Disposed;
        event EventHandler FileOk;
        event EventHandler HelpRequest;

        bool AddExtension { get; set; }
        bool AutoUpgradeEnabled { get; set; }
        bool CheckFileExists { get; set; }
        bool CheckPathExists { get; set; }
        string DefaultExt { get; set; }
        bool DereferenceLinks { get; set; }
        string FileName { get; set; }
        string Filter { get; set; }
        int FilterIndex { get; set; }
        string InitialDirectory { get; set; }
        bool Multiselect { get; set; }
        bool ReadOnlyChecked { get; set; }
        bool RestoreDirectory { get; set; }
        bool ShowHelp { get; set; }
        bool ShowReadOnly { get; set; }
        ISite Site { get; set; }
        bool SupportMultiDottedExtensions { get; set; }
        object Tag { get; set; }
        string Title { get; set; }
        bool ValidateNames { get; set; }

        IContainer Container { get; }
        FileDialogCustomPlacesCollection CustomPlaces { get; }
        string[] FileNames { get; }
        string SafeFileName { get; }
        string[] SafeFileNames { get; }

        ObjRef CreateObjRef(Type requestedType);
        object GetLifetimeService();
        object InitializeLifetimeService();
        Stream OpenFile();
        void Reset();
        DialogResult ShowDialog();
    }

    public class OpenFileDialog : IOpenFileDialog
    {
        private readonly System.Windows.Forms.OpenFileDialog _openFileDialog;

        public OpenFileDialog()
        {
            _openFileDialog = new System.Windows.Forms.OpenFileDialog();
            _openFileDialog.Disposed += OpenFileDialog_Disposed;
            _openFileDialog.FileOk += OpenFileDialog_FileOk;
            _openFileDialog.HelpRequest += OpenFileDialog_HelpRequest;
        }

        public bool AddExtension
        {
            get => _openFileDialog.AddExtension;
            set => _openFileDialog.AddExtension = value;
        }

        public bool AutoUpgradeEnabled
        {
            get => _openFileDialog.AutoUpgradeEnabled;
            set => _openFileDialog.AutoUpgradeEnabled = value;
        }

        public bool CheckFileExists
        {
            get => _openFileDialog.CheckFileExists;
            set => _openFileDialog.CheckFileExists = value;
        }

        public bool CheckPathExists
        {
            get => _openFileDialog.CheckPathExists;
            set => _openFileDialog.CheckPathExists = value;
        }

        public IContainer Container => _openFileDialog.Container;

        public ObjRef CreateObjRef(Type requestedType)
        {
            return _openFileDialog.CreateObjRef(requestedType);
        }

        public FileDialogCustomPlacesCollection CustomPlaces => _openFileDialog.CustomPlaces;

        public string DefaultExt
        {
            get => _openFileDialog.DefaultExt;
            set => _openFileDialog.DefaultExt = value;
        }

        public bool DereferenceLinks
        {
            get => _openFileDialog.DereferenceLinks;
            set => _openFileDialog.DereferenceLinks = value;
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

            _openFileDialog.Dispose();
            _isDisposed = true;
        }

        public event EventHandler Disposed;
        private void OpenFileDialog_Disposed(object sender, EventArgs e)
        {
            Disposed?.Invoke(sender, e);
        }

        public override bool Equals(object obj)
        {
            return _openFileDialog.Equals(obj);
        }

        public string FileName
        {
            get => _openFileDialog.FileName;
            set => _openFileDialog.FileName = value;
        }

        public string[] FileNames => _openFileDialog.FileNames;

        public event EventHandler FileOk;
        private void OpenFileDialog_FileOk(object sender, EventArgs e)
        {
            FileOk?.Invoke(sender, e);
        }

        public string Filter
        {
            get => _openFileDialog.Filter;
            set => _openFileDialog.Filter = value;
        }

        public int FilterIndex
        {
            get => _openFileDialog.FilterIndex;
            set => _openFileDialog.FilterIndex = value;
        }

        public override int GetHashCode()
        {
            return _openFileDialog.GetHashCode();
        }

        public object GetLifetimeService()
        {
            return _openFileDialog.GetLifetimeService();
        }

        public event EventHandler HelpRequest;
        private void OpenFileDialog_HelpRequest(object sender, EventArgs e)
        {
            HelpRequest?.Invoke(sender, e);
        }

        public string InitialDirectory
        {
            get => _openFileDialog.InitialDirectory;
            set => _openFileDialog.InitialDirectory = value;
        }

        public object InitializeLifetimeService()
        {
            return _openFileDialog.InitializeLifetimeService();
        }

        public bool Multiselect
        {
            get => _openFileDialog.Multiselect;
            set => _openFileDialog.Multiselect = value;
        }

        public Stream OpenFile()
        {
            return _openFileDialog.OpenFile();
        }

        public bool ReadOnlyChecked
        {
            get => _openFileDialog.ReadOnlyChecked;
            set => _openFileDialog.ReadOnlyChecked = value;
        }

        public void Reset()
        {
            _openFileDialog.Reset();
        }

        public bool RestoreDirectory
        {
            get => _openFileDialog.RestoreDirectory;
            set => _openFileDialog.RestoreDirectory = value;
        }

        public string SafeFileName => _openFileDialog.SafeFileName;

        public string[] SafeFileNames => _openFileDialog.SafeFileNames;

        public DialogResult ShowDialog()
        {
            return _openFileDialog.ShowDialog();
        }

        public bool ShowHelp
        {
            get => _openFileDialog.ShowHelp;
            set => _openFileDialog.ShowHelp = value;
        }

        public bool ShowReadOnly
        {
            get => _openFileDialog.ShowReadOnly;
            set => _openFileDialog.ShowReadOnly = value;
        }

        public ISite Site
        {
            get => _openFileDialog.Site;
            set => _openFileDialog.Site = value;
        }

        public bool SupportMultiDottedExtensions
        {
            get => _openFileDialog.SupportMultiDottedExtensions;
            set => _openFileDialog.SupportMultiDottedExtensions = value;
        }

        public object Tag
        {
            get => _openFileDialog.Tag;
            set => _openFileDialog.Tag = value;
        }

        public string Title
        {
            get => _openFileDialog.Title;
            set => _openFileDialog.Title = value;
        }

        public override string ToString()
        {
            return _openFileDialog.ToString();
        }

        public bool ValidateNames
        {
            get => _openFileDialog.ValidateNames;
            set => _openFileDialog.ValidateNames = value;
        }
    }
}

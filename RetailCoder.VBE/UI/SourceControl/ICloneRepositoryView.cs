using System;

namespace Rubberduck.UI.SourceControl
{
    public interface ICloneRepositoryView
    {
        IFolderBrowserFactory FolderBrowserFactory { get; set; }

        string RemotePath { get; set; }
        bool IsValidRemotePath { get; set; }

        string LocalDirectory { get; set; }

        event EventHandler<CloneRepositoryEventArgs> Confirm;
        event EventHandler<EventArgs> Cancel;
        event EventHandler<EventArgs> RemotePathChanged;

        void Show();
        void Hide();
        void Close();
    }

    public class CloneRepositoryEventArgs : EventArgs
    {
        public string RemotePath { get; private set; }
        public string LocalDirectory { get; private set; }

        public CloneRepositoryEventArgs(string remotePath, string localDirectory)
        {
            RemotePath = remotePath;
            LocalDirectory = localDirectory;
        }
    }
}
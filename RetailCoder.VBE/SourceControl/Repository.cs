using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace Rubberduck.SourceControl
{
    [ComVisible(true)]
    [Guid("E8509738-3A06-4E8F-85FE-16F63F5A6DC3")]
    public interface IRepository
    {
        [DispId(0)]
        string Name { get; }
        
        [DispId(1)]
        [Description("FilePath of local repository.")]
        string LocalLocation { get; }

        [DispId(2)]
        [Description("FilePath or URL of remote repository.")]
        string RemoteLocation { get; }
    }

    [ComVisible(true)]
    [Guid("B2965961-7240-40CD-BE16-9425E2FB003C")]
    [ProgId("Rubberduck.Repository")]
    [ClassInterface(ClassInterfaceType.None)]
    public class Repository : IRepository
    {
        public string Name { get; private set; }
        public string LocalLocation { get; private set; }
        public string RemoteLocation { get; private set; }

        public Repository(string name, string localDirectory, string remotePathOrUrl)
        {
            this.Name = name;
            this.LocalLocation = localDirectory;
            this.RemoteLocation = remotePathOrUrl;
        }
    }
}

using System.Runtime.InteropServices;

namespace Rubberduck.SourceControl
{
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
            Name = name;
            LocalLocation = localDirectory;
            RemoteLocation = remotePathOrUrl;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.SourceControl
{
    public interface IRepository
    {
        string LocalLocation { get; }
        string Name { get; }
        string RemoteLocation { get; }
    }

    public class Repository : Rubberduck.SourceControl.IRepository
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

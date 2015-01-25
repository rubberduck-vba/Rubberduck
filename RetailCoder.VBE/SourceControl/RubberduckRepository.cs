using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.SourceControl
{
    public class Repository
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

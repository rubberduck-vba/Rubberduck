using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

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
}

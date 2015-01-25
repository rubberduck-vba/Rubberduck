using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Microsoft.Vbe.Interop;

//todo: implement source control base
namespace Rubberduck.SourceControl
{
    internal class FileMediator
    {
        private VBProject project;
        public FileMediator(VBProject project)
        {
            this.project = project;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace RubberduckTests.CommonObjects
{
    /// <summary>Helps solving problem with Mock which I was unable to set up to support <code>Cast</code> extension</summary>
    internal class VbProjecstFake :VBProjects
    {
        private IEnumerable<VBProject> _projectList = null; 

        internal VbProjecstFake(IEnumerable<VBProject> projectList )
        {
            _projectList = projectList;
        }
        
        public int Count
        {
            get { return _projectList != null ? _projectList.ToList().Count : 0;}
        }

        public System.Collections.IEnumerator GetEnumerator()
        {
            if (_projectList == null) throw new ArgumentNullException();
            return _projectList.GetEnumerator();
        }

        public VBProject Item(object index)
        {
            if (_projectList == null) throw new ArgumentNullException();
            if (index is int)
            {
                return _projectList.ToList()[(int) index];
            }
            else if (index is VBProject)
            {
                return _projectList.ToList().Find(x => x.Name == ((VBProject) index).Name);
            }
            else
            {
                throw new ArgumentException("Type is not supported");
            }
        }

        public VBProject Add(vbext_ProjectType Type) { throw new NotImplementedException(); }

        public VBProject Open(string bstrPath) { throw new NotImplementedException(); }

        public VBE Parent { get { throw new NotImplementedException(); }}

        public void Remove(VBProject lpc) {throw new NotImplementedException();}

        public VBE VBE { get { throw new NotImplementedException(); }
        }
    }
}

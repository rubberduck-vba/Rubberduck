using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Vbe.Interop;

namespace RubberduckTests.CommonObjects
{
    internal class VbComponentsFake : VBComponents
    {
        private IEnumerable<VBComponent> _componentList = null;

        internal VbComponentsFake(IEnumerable<VBComponent> componentList)
        {
            _componentList = componentList;
        }

        public int Count
        {
            get { return _componentList != null ? _componentList.ToList().Count : 0; }
        }

        public System.Collections.IEnumerator GetEnumerator()
        {
            if (_componentList == null) throw new ArgumentNullException();
            return _componentList.GetEnumerator();
        }

        public VBComponent Item(object index)
        {
            if (_componentList == null) throw new ArgumentNullException();
            if (index is int)
            {
                return _componentList.ToList()[(int)index];
            }
            else if (index is VBProject)
            {
                return _componentList.ToList().Find(x => x.Name == ((VBProject)index).Name);
            }
            else
            {
                throw new ArgumentException("Type is not supported");
            }
        }
        
        public VBComponent Add(vbext_ComponentType ComponentType) { throw new NotImplementedException();}

        public VBComponent AddCustom(string ProgId) { throw new NotImplementedException();}

        public VBComponent AddMTDesigner(int index = 0) { throw new NotImplementedException(); }
        
        public VBComponent Import(string FileName) { throw new NotImplementedException(); }
        
        public VBProject Parent { get { throw new NotImplementedException(); } }

        public void Remove(VBComponent VBComponent) { throw new NotImplementedException();}

        public VBE VBE { get { throw new NotImplementedException(); }}
    }
}

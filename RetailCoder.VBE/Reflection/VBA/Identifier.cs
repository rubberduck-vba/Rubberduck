using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RetailCoderVBE.Reflection.VBA
{
    internal class Identifier
    {
        private readonly string _name;
        public string Name { get { return _name; } }

        private readonly string _typeName;
        public string TypeName { get { return _typeName; } }

        public Identifier(string name, string typeName)
        {
            _name = name;
            _typeName = typeName;
        }
    }
}

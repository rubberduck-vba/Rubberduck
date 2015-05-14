using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.Refactorings.ReorderParameters
{
    public class Parameter
    {
        public string Name { get; set; }

        public Parameter(string name)
        {
            Name = name;
        }
    }
}

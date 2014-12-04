using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Xml.Serialization;

namespace Rubberduck.Config
{
    [ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class CodeInspectionSettings
    {
        [XmlArrayItemAttribute("CodeInspection", IsNullable = false)]
        public CodeInspection[] CodeInspections { get; set; }
    }

    [ComVisible(false)]
    [XmlTypeAttribute(AnonymousType = true)]
    public class CodeInspection
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlAttribute]
        public bool On { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Rubberduck.Inspections;
using Antlr4.Runtime;

namespace Rubberduck.UI
{
    [System.Runtime.InteropServices.ComVisible(false)]
    public class Instruction
    {
        public QualifiedModuleName QualifiedModuleName { get; private set; }
        public ParserRuleContext ParserRuleContext { get; private set; }

        public Instruction(QualifiedModuleName qualifiedModuleName, ParserRuleContext context)
        {
            this.QualifiedModuleName = qualifiedModuleName;
            this.ParserRuleContext = context;
        }

    }
}

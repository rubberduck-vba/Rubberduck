using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Excel;

namespace Rubberduck.VBA
{
    [ComVisible(false)]
    public class CodeBlockNode : SyntaxTreeNode
    {
        public CodeBlockNode(Instruction instruction, string scope, Match match, string[] endingMarkers, Type childSyntaxType, IEnumerable<SyntaxTreeNode> nodes)
            : base(instruction, scope, match, nodes)
        {
            _endingMarkers = endingMarkers;
            _childSyntaxType = childSyntaxType;
        }

        private readonly string[] _endingMarkers;
        public IEnumerable<string> EndOfBlockMarkers { get { return _endingMarkers; } }

        private readonly Type _childSyntaxType;
        public Type ChildSyntaxType { get { return _childSyntaxType; } }
    }
}

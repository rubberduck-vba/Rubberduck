using System;
using System.Runtime.InteropServices;

namespace Rubberduck.VBA.Parser.Grammar
{
    [ComVisible(false)]
    [Flags]
    public enum SyntaxType
    {
        Syntax = 0,
        /// <summary>
        /// Indicates that this syntax produces child nodes.
        /// </summary>
        HasChildNodes = 1,
        /// <summary>
        /// Indicates that this syntax isn't part of the language's general grammar, 
        /// e.g. 
        /// </summary>
        IsChildNodeSyntax = 2
    }
}
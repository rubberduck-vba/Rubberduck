using Rubberduck.Parsing.Grammar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ResultVariableDeclaration : Declaration
    {
        public ResultVariableDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration function,
            string asTypeName,
            string typeHint,
            bool isArray)
            : base(
                  qualifiedName,
                  function,
                  function,
                  asTypeName,
                  typeHint,
                  false,
                  false,
                  Accessibility.Implicit,
                  DeclarationType.ResultVariable,
                  null,
                  Selection.Home,
                  isArray,
                  null,
                  false)
        {
            Function = function;
        }

        public Declaration Function { get; }
    }
}

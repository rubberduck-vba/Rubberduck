using Rubberduck.Parsing.Symbols;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.Exceptions
{
    public class MoveMemberUnsupportedMoveException : RefactoringException
    {
        public MoveMemberUnsupportedMoveException() { }

        public MoveMemberUnsupportedMoveException(Declaration declaration)
        {
            TargetDeclaration = declaration;
        }
        public Declaration TargetDeclaration { get; }
    }
}

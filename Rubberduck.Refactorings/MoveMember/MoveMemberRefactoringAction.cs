using Microsoft.CSharp.RuntimeBinder;
using NLog;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Exceptions;
using Rubberduck.VBEditor.ComManagement;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.Utility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.Refactorings.MoveMember
{
    public class MoveMemberRefactoringAction : IRefactoringAction<MoveMemberModel>
    {
        private readonly IRefactoringAction<MoveMemberModel> _moveMemberToNewModuleRefactoringAction;
        private readonly IRefactoringAction<MoveMemberModel> _moveMemberToExistingModuleRefactoringAction;

        public MoveMemberRefactoringAction(MoveMemberToNewModuleRefactoring moveMemberToNewModuleRefactoring, 
                                            MoveMemberToExistingModuleRefactoring moveMemberToExistingModuleRefactoring)
        {
            _moveMemberToNewModuleRefactoringAction = moveMemberToNewModuleRefactoring;
            _moveMemberToExistingModuleRefactoringAction = moveMemberToExistingModuleRefactoring;
        }

        public void Refactor(MoveMemberModel model)
        {
            if (model.Destination.IsExistingModule(out _))
            {
                _moveMemberToExistingModuleRefactoringAction.Refactor(model);
                return;
            }

            _moveMemberToNewModuleRefactoringAction.Refactor(model);
        }
    }
}

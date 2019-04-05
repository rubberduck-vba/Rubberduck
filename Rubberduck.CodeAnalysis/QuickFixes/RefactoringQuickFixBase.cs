using System;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Rewriter;
using Rubberduck.Refactorings;
using Rubberduck.Refactorings.Exceptions;

namespace Rubberduck.Inspections.QuickFixes
{
    public abstract class RefactoringQuickFixBase : QuickFixBase
    {
        protected readonly IRefactoring Refactoring;

        protected RefactoringQuickFixBase(IRefactoring refactoring, params Type[] inspections)
            : base(inspections)
        {
            Refactoring = refactoring;
        }

        //The rewriteSession is optional since it is not used in refactoring quickfixes.
        public override void Fix(IInspectionResult result, IRewriteSession rewriteSession = null)
        {
            try
            {
                Refactor(result);
            }
            catch (RefactoringAbortedException)
            {}
            catch (RewriteFailedException)
            {
                //We rethrow because this information is required by the QuickFixProvider to trigger the failure notiication. 
                throw;
            }
            catch (RefactoringException exception)
            {
                //This is an error: the inspection returned an invalid result. 
                Logger.Error(exception);
            }
        }

        protected abstract void Refactor(IInspectionResult result);

        public override bool CanFixInProcedure => false;
        public override bool CanFixInModule => false;
        public override bool CanFixInProject => false;
    }
}
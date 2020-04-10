
namespace Rubberduck.Refactorings.Exceptions
{
    public class MoveMemberUnsupportedMoveException : RefactoringException
    {
        public MoveMemberUnsupportedMoveException() { }

        public MoveMemberUnsupportedMoveException(string message)
        {
            Message = message;
        }
        public override string Message { get; }
    }
}

using System.Collections.Generic;
using Parameter = Rubberduck.Refactorings.RemoveParameters.Parameter;

namespace Rubberduck.Refactorings.Exceptions.RemoveParameter
{
    public class MultipleParametersSelectedException : RefactoringException
    {
        public MultipleParametersSelectedException(ICollection<Parameter> selectedParameters)
        {
            SelectedParameters = selectedParameters;
        }

        public ICollection<Parameter> SelectedParameters { get; }
    }
}
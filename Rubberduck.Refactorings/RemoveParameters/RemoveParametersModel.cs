using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.RemoveParameters
{
    public class RemoveParametersModel : IRefactoringModel
    {
        public RemoveParametersModel(Declaration target)
        {
            Parameters = new List<Parameter>();
            RemoveParameters = new List<Parameter>();

            OriginalTarget = target;
            TargetDeclaration = target;
            LoadParameters();
        }

        private Declaration _target; 
        public Declaration TargetDeclaration
        {
            get => _target;
            set
            {
                if (value == null)
                {
                    return;
                }

                _target = value;
                LoadParameters();
            }
        }

        public Declaration OriginalTarget { get; }
        public List<Parameter> Parameters { get; }
        public List<Parameter> RemoveParameters { get; set; }

        public bool IsInterfaceMemberRefactoring { get; set; }
        public bool IsEventRefactoring { get; set; }
        public bool IsPropertyRefactoringWithGetter { get; set; }

        private void LoadParameters()
        {
            Parameters.Clear();
            RemoveParameters.Clear();

            if (!(TargetDeclaration is IParameterizedDeclaration parameterizedDeclaration))
            {
                return;
            }
            
            Parameters.AddRange(parameterizedDeclaration.Parameters.Select(arg => new Parameter(arg)));

            if (TargetDeclaration.DeclarationType == DeclarationType.PropertyLet
                || TargetDeclaration.DeclarationType == DeclarationType.PropertySet)
            {
                Parameters.Remove(Parameters.Last());
            }
        }
    }
}

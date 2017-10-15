using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.RemoveParameters
{
    /// <summary>
    /// Contains data about a method parameter for the Remove Parameters refactoring.
    /// </summary>
    public class Parameter : ViewModelBase
    {
        public ParameterDeclaration Declaration { get; }
        public string Name { get; }
        public bool IsParamArray => Declaration.IsParamArray;

        private bool _isRemoved;
        public bool IsRemoved
        {
            get { return _isRemoved; }
            set
            {
                _isRemoved = value;
                OnPropertyChanged();
            }
        }

        public Parameter(Declaration declaration, bool isRemoved = false)
        {
            Declaration = (ParameterDeclaration)declaration;
            Name = declaration.Context.GetText().RemoveExtraSpacesLeavingIndentation();
            IsRemoved = isRemoved;
        }
    }
}

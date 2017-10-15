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
        public Declaration Declaration { get; private set; }
        public string Name { get; private set; }
        public int Index { get; private set; }

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

        public Parameter(Declaration declaration, int index, bool isRemoved = false)
        {
            Declaration = declaration;
            Name = declaration.Context.GetText().RemoveExtraSpacesLeavingIndentation();
            Index = index;
            IsRemoved = isRemoved;
        }
    }
}

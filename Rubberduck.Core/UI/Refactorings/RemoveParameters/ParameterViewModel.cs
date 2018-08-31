using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.RemoveParameters;

namespace Rubberduck.UI.Refactorings.RemoveParameters
{
    public class ParameterViewModel : ViewModelBase
    {
        private Parameter wrapped;
        internal Parameter Wrapped { get => wrapped; }

        public ParameterViewModel(Parameter wrapped)
        {
            this.wrapped = wrapped;
        }

        private bool _isRemoved;
        public bool IsRemoved
        {
            get => _isRemoved;
            set
            {
                _isRemoved = value;
                OnPropertyChanged();
            }
        }

        public string Name { get => wrapped.Name; }
        public Declaration Declaration { get => wrapped.Declaration; }

    }

    internal static class ConversionExtensions
    {
        public static Parameter ToModel(this ParameterViewModel viewModel)
        {
            return viewModel.Wrapped;
        }

        public static ParameterViewModel ToViewModel(this Parameter model)
        {
            return new ParameterViewModel(model);
        }
    }
}

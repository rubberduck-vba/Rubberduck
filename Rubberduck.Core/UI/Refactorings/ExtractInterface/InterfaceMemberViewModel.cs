using Rubberduck.Refactorings.ExtractInterface;

namespace Rubberduck.UI.Refactorings.ExtractInterface
{
    internal class InterfaceMemberViewModel : ViewModelBase 
    {
        public InterfaceMemberViewModel(InterfaceMember model)
        {
            Wrapped = model;
        }

        internal InterfaceMember Wrapped { get; }

        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public string FullMemberSignature => Wrapped?.FullMemberSignature;
    }

    internal static class ConversionExtensions
    {
        public static InterfaceMember ToModel(this InterfaceMemberViewModel viewModel)
        {
            return viewModel.Wrapped;
        }

        public static InterfaceMemberViewModel ToViewModel(this InterfaceMember model)
        {
            return new InterfaceMemberViewModel(model);
        }
    }
}

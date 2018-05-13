using Rubberduck.Refactorings.ExtractInterface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.Refactorings.ExtractInterface
{
    internal class InterfaceMemberViewModel : ViewModelBase 
    {
        private readonly InterfaceMember _wrapped;
        internal InterfaceMember Wrapped { get => _wrapped; }


        private bool _isSelected;
        private InterfaceMember model;

        public InterfaceMemberViewModel(InterfaceMember model)
        {
            this.model = model;
        }

        public bool IsSelected
        {
            get => _isSelected;
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public string FullMemberSignature { get => _wrapped.FullMemberSignature; }
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

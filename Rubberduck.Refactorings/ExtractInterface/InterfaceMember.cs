using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class InterfaceMember : INotifyPropertyChanged
    {
        private readonly ModuleBodyElementDeclaration _element;
        private readonly ICodeBuilder _codeBuilder;

        public InterfaceMember(Declaration member, ICodeBuilder codeBuilder)
        {
            Member = member;
            if (!(member is ModuleBodyElementDeclaration mbed))
            {
                throw new ArgumentException();
            }
            _element = mbed;
            _codeBuilder = codeBuilder;
        }

        public Declaration Member { get; }

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

        public string FullMemberSignature => _codeBuilder.ImprovedFullMemberSignature(_element);

        public string Body => _codeBuilder.BuildMemberBlockFromPrototype(_element);

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
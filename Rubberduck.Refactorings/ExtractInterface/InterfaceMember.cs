using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Refactorings.Common;
using Rubberduck.Refactorings.ImplementInterface;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class InterfaceMember : INotifyPropertyChanged
    {
        private readonly ModuleBodyElementDeclaration _element;

        public InterfaceMember(Declaration member)
        {
            Member = member;
            if (!(member is ModuleBodyElementDeclaration mbed))
            {
                throw new ArgumentException();
            }
            _element = mbed;
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

        public string FullMemberSignature => _element.FullMemberSignature();

        public string Body => _element.AsCodeBlock();

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
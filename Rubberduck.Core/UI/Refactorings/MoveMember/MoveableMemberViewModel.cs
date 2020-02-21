using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;
using Rubberduck.Refactorings.MoveMember;
using Rubberduck.Refactorings.MoveMember.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Rubberduck.UI.Refactorings.MoveMember
{
    public class MoveableMemberSetViewModel : INotifyPropertyChanged
    {
        private IMoveableMemberSet _moveable;
        private MoveMemberViewModel _parentModel;

        public MoveableMemberSetViewModel(MoveMemberViewModel vmodel, IMoveableMemberSet moveable)
        {
            _parentModel = vmodel;
            _moveable = moveable;
        }

        public string IdentifierName => _moveable.IdentifierName;

        public string MovedIdentiferName => _moveable.MovedIdentifierName;

        public bool IsSelected
        {
            get => _moveable.IsSelected;
            set
            {
                _moveable.IsSelected = value;
                OnPropertyChanged();
                _parentModel.RefreshPreview(this);
            }
        }

        private string _displaySignature;
        public string MemberDisplaySignature
        {
            get
            {
                if (_displaySignature != null)
                {
                    return _displaySignature;
                }

                _displaySignature = string.Empty;

                var accessibility = _moveable.Member.Accessibility == Accessibility.Implicit
                    ? Tokens.Public
                    : _moveable.Member.Accessibility.TokenString();

                if (_moveable.Member is ModuleBodyElementDeclaration mbed)
                {
                    _displaySignature = mbed.FullyDefinedSignature();
                    if (_moveable.Member.DeclarationType.HasFlag(DeclarationType.Property))
                    {
                        var LetSetGetTokens = new SortedDictionary<string, string>();
                        foreach (var member in _moveable.Members)
                        {
                            if (member.DeclarationType.Equals(DeclarationType.PropertyLet)) { LetSetGetTokens.Add("a", Tokens.Let); }
                            if (member.DeclarationType.Equals(DeclarationType.PropertySet)) { LetSetGetTokens.Add("b", Tokens.Set); }
                            if (member.DeclarationType.Equals(DeclarationType.PropertyGet)) { LetSetGetTokens.Add("c", Tokens.Get); }
                        }

                        if (mbed.DeclarationType.Equals(DeclarationType.PropertyLet))
                        {
                            _displaySignature = _displaySignature.Replace($"{Tokens.Property} {Tokens.Let}", $"{Tokens.Property} {string.Join("\\", LetSetGetTokens.Values)}");
                        }
                        else if (mbed.DeclarationType.Equals(DeclarationType.PropertySet))
                        {
                            _displaySignature = _displaySignature.Replace($"{Tokens.Property} {Tokens.Set}", $"{Tokens.Property} {string.Join("\\", LetSetGetTokens.Values)}");
                        }
                        else
                        {
                            _displaySignature = _displaySignature.Replace($"{Tokens.Property} {Tokens.Get}", $"{Tokens.Property} {string.Join("\\", LetSetGetTokens.Values)}");
                        }
                    }
                }
                else if (_moveable.Member.IsConstant())
                {
                    _displaySignature = $"{accessibility} {_moveable.Member.IdentifierName} {Tokens.Const}";
                }
                else if (_moveable.Member.IsField())
                {
                    _displaySignature = $"{accessibility} {_moveable.Member.IdentifierName}";
                }

                _displaySignature = _moveable.Member.AsTypeName == null ? _displaySignature : $"{_displaySignature} {Tokens.As} {_moveable.Member.AsTypeName}";
                return _displaySignature;
            }
        }

        public override int GetHashCode()
        {
            return _moveable.IdentifierName.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (obj is MoveableMemberSet moveable)
            {
                return moveable.IdentifierName == _moveable.IdentifierName && moveable.MovedIdentifierName == _moveable.MovedIdentifierName;
            }
            return false;
        }


        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}

using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Properties;
using Rubberduck.Refactorings.Common;
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
            ToDisplayString = BuildDisplaySignature(_moveable);
        }

        public string IdentifierName => _moveable.IdentifierName;

        public string MovedIdentiferName => _moveable.MovedIdentifierName;

        public bool IsPublicMember => _moveable.Members.Any(m => m.IsMember() && !m.HasPrivateAccessibility());

        public bool IsPrivateMember => !IsPublicMember && _moveable.Member.IsMember();

        public bool IsPublicConstant => _moveable.Member.IsModuleConstant() && !_moveable.Member.HasPrivateAccessibility();

        public bool IsPrivateConstant => _moveable.Member.IsModuleConstant() && _moveable.Member.HasPrivateAccessibility();

        public bool IsPublicField => _moveable.Member.IsField() && !_moveable.Member.HasPrivateAccessibility();

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

        public string ToDisplayString { private set; get; }

        public static string BuildDisplaySignature(IMoveableMemberSet moveable)
        {
            var displaySignature = string.Empty;
            var accessibility = moveable.Member.Accessibility == Accessibility.Implicit
                ? Tokens.Public
                : moveable.Member.Accessibility.TokenString();

            if (moveable.Member is ModuleBodyElementDeclaration moduleBodyElementDeclaration)
            {
                displaySignature = moduleBodyElementDeclaration.FullMemberSignature();

                if (moveable.Member.DeclarationType.HasFlag(DeclarationType.Property))
                {
                    //Use the Property Get signature as the default if available
                    var propertyTemplate = moveable.Members.SingleOrDefault(mm => mm.DeclarationType.Equals(DeclarationType.PropertyGet))
                                            ?? moveable.Member;

                    displaySignature = moveable.Member.Equals(propertyTemplate)
                        ? displaySignature
                        : (propertyTemplate as ModuleBodyElementDeclaration).FullMemberSignature();

                    //Force the Tokens order to be Let\Set\Get so that 'Get' is closest to the signature
                    var LetSetGetTokens = new SortedDictionary<string, string>();
                    foreach (var member in moveable.Members)
                    {
                        switch (member.DeclarationType)
                        {
                            case DeclarationType.PropertyLet:
                                LetSetGetTokens.Add("a", Tokens.Let);
                                break;
                            case DeclarationType.PropertySet:
                                LetSetGetTokens.Add("b", Tokens.Set);
                                break;
                            case DeclarationType.PropertyGet:
                                LetSetGetTokens.Add("c", Tokens.Get);
                                break;
                            default:
                                throw new ArgumentException();
                        }
                    }

                    var displaySignaturePrefix = $"{Tokens.Property} {string.Join("\\", LetSetGetTokens.Values)}";
                    switch (propertyTemplate.DeclarationType)
                    {
                        case DeclarationType.PropertyGet:
                            displaySignature = displaySignature.Replace($"{Tokens.Property} {Tokens.Get}", displaySignaturePrefix);
                            break;
                        case DeclarationType.PropertyLet:
                            displaySignature = displaySignature.Replace($"{Tokens.Property} {Tokens.Let}", displaySignaturePrefix);
                            break;
                        case DeclarationType.PropertySet:
                            displaySignature = displaySignature.Replace($"{Tokens.Property} {Tokens.Set}", displaySignaturePrefix);
                            break;
                        default:
                            throw new ArgumentException();
                    }
                }
            }
            else if (moveable.Member.IsConstant())
            {
                displaySignature = $"{accessibility} {Tokens.Const} {moveable.Member.IdentifierName}";
                var constValue = string.Empty;
                if (moveable.Member.Context.TryGetChildContext<VBAParser.LiteralExprContext>(out var litExpr))
                {
                    constValue = litExpr.GetText();
                }
                else if (moveable.Member.Context.TryGetChildContext<VBAParser.LExprContext>(out var lExpr))
                {
                    constValue = lExpr.GetText();
                }
                displaySignature = moveable.Member.AsTypeName == null ? displaySignature : $"{displaySignature} {Tokens.As} {moveable.Member.AsTypeName} = {constValue}";
            }
            else if (moveable.Member.IsField())
            {
                displaySignature = $"{accessibility} {moveable.Member.IdentifierName}";
                displaySignature = moveable.Member.AsTypeName == null ? displaySignature : $"{displaySignature} {Tokens.As} {moveable.Member.AsTypeName}";
            }
            return displaySignature;
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

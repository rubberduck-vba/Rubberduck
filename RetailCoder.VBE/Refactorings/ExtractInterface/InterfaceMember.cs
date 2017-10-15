using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.UI;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class Parameter
    {
        public string ParamAccessibility { get; set; }
        public string ParamName { get; set; }
        public string ParamType { get; set; }

        public override string ToString()
        {
            return ParamAccessibility + " " + ParamName + " As " + ParamType;
        }
    }

    public class InterfaceMember : ViewModelBase
    {
        public Declaration Member { get; }
        public IEnumerable<Parameter> MemberParams { get; }
        private string Type { get; }

        private string MemberType { get; set; }

        private bool _isSelected;
        public bool IsSelected
        {
            get { return _isSelected; }
            set
            {
                _isSelected = value;
                OnPropertyChanged();
            }
        }

        public string Identifier { get; }

        public string FullMemberSignature
        {
            get
            {
                var signature = MemberType + " " + Member.IdentifierName + "(" +
                    string.Join(", ", MemberParams) + ")";

                return Type == null ? signature : signature + " As " + Type;
            }
        }

        public InterfaceMember(Declaration member)
        {
            Member = member;
            Identifier = member.IdentifierName;
            Type = member.AsTypeName;
            
            GetMethodType();

            var memberWithParams = member as IParameterizedDeclaration;
            if (memberWithParams != null)
            {
                MemberParams = memberWithParams.Parameters
                    .OrderBy(o => o.Selection.StartLine)
                    .ThenBy(t => t.Selection.StartColumn)
                    .Select(p => new Parameter
                    {
                        ParamAccessibility =
                            ((VBAParser.ArgContext) p.Context).BYVAL() != null ? Tokens.ByVal : Tokens.ByRef,
                        ParamName = p.IdentifierName,
                        ParamType = p.AsTypeName
                    })
                    .ToList();
            }
            else
            {
                MemberParams = new List<Parameter>();
            }

            if (MemberType == "Property Get")
            {
                MemberParams = MemberParams.Take(MemberParams.Count() - 1);
            }

            IsSelected = false;
        }

        private void GetMethodType()
        {
            var context = Member.Context;

            var subStmtContext = context as VBAParser.SubStmtContext;
            if (subStmtContext != null)
            {
                MemberType = Tokens.Sub;
            }

            var functionStmtContext = context as VBAParser.FunctionStmtContext;
            if (functionStmtContext != null)
            {
                MemberType = Tokens.Function;
            }

            var propertyGetStmtContext = context as VBAParser.PropertyGetStmtContext;
            if (propertyGetStmtContext != null)
            {
                MemberType = Tokens.Property + " " + Tokens.Get;
            }

            var propertyLetStmtContext = context as VBAParser.PropertyLetStmtContext;
            if (propertyLetStmtContext != null)
            {
                MemberType = Tokens.Property + " " + Tokens.Let;
            }

            var propertySetStmtContext = context as VBAParser.PropertySetStmtContext;
            if (propertySetStmtContext != null)
            {
                MemberType = Tokens.Property + " " + Tokens.Set;
            }
        }

        public string Body => "Public " + FullMemberSignature + Environment.NewLine +
                              "End " + MemberType.Split(' ').First() + Environment.NewLine;
    }
}

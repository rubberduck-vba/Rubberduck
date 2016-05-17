using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;

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

    public class InterfaceMember
    {
        private Declaration Member { get; set; }
        private IEnumerable<Parameter> MemberParams { get; set; }
        private string Type { get; set; }

        private string MemberType { get; set; }

        public bool IsSelected { get; set; }

        public string FullMemberSignature
        {
            get
            {
                var signature = MemberType + " " + Member.IdentifierName + "(" +
                    string.Join(", ", MemberParams) + ")";

                return Type == null ? signature : signature + " As " + Type;
            }
        }

        public InterfaceMember(Declaration member, IEnumerable<Declaration> declarations)
        {
            Member = member;
            Type = member.AsTypeName;
            
            GetMethodType();

            MemberParams = declarations.Where(item => item.DeclarationType == DeclarationType.Parameter &&
                                          item.ParentScope == Member.Scope)
                                       .OrderBy(o => o.Selection.StartLine)
                                       .ThenBy(t => t.Selection.StartColumn)
                                       .Select(p => new Parameter
                                       {
                                           ParamAccessibility = ((VBAParser.ArgContext)p.Context).BYREF() == null ? Tokens.ByVal : Tokens.ByRef,
                                           ParamName = p.IdentifierName,
                                           ParamType = p.AsTypeName
                                       })
                                       .ToList();

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

        public string Body
        {
            get
            {
                return "Public " + FullMemberSignature + Environment.NewLine +
                "End " + MemberType.Split(' ').First() + Environment.NewLine;
            }
        }
    }
}

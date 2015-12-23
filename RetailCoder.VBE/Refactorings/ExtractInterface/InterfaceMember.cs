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
        private string PropertyType { get; set; }

        public bool IsSelected { get; set; }
        public string MemberSignature
        {
            get
            {
                var signature = Member.IdentifierName + "(" +
                    string.Join(", ", MemberParams.Select(m => m.ParamType)) + ")";

                return Type == null ? signature : signature + " As " + Type;
            }
        }

        public string FullMemberSignature
        {
            get
            {
                var signature = Member.IdentifierName + "(" +
                    string.Join(", ", MemberParams) + ")";

                return Type == null ? signature : signature + " As " + Type;
            }
        }

        public InterfaceMember(Declaration member, IEnumerable<Declaration> declarations)
        {
            Member = member;
            Type = member.AsTypeName;

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

            GetMethodType();

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
                MemberType = Tokens.Property;
                PropertyType = Tokens.Get;
            }

            var propertyLetStmtContext = context as VBAParser.PropertyLetStmtContext;
            if (propertyLetStmtContext != null)
            {
                MemberType = Tokens.Property;
                PropertyType = Tokens.Let;
            }

            var propertySetStmtContext = context as VBAParser.PropertySetStmtContext;
            if (propertySetStmtContext != null)
            {
                MemberType = Tokens.Property;
                PropertyType = Tokens.Set;
            }
        }

        public override string ToString()
        {
            return "Public " + MemberType + " " + PropertyType + " " + FullMemberSignature + Environment.NewLine + "End " + MemberType +
                   Environment.NewLine;
        }
    }
}
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Refactorings.ExtractInterface
{
    public class Parameter
    {
        public Accessibility ParamAccessibility { get; set; }
        public string ParamName { get; set; }
        public string ParamType { get; set; }

        public override string ToString()
        {
            return ParamAccessibility + " " + ParamName + " As " + ParamType;
        }
    }

    public class InterfaceMember
    {
        public Declaration Member { get; set; }
        public IEnumerable<Parameter> MemberParams { get; set; }
        public bool IsSelected { get; set; }
        public string Signature { get { return ToString(); } }

        public InterfaceMember(Declaration member, IEnumerable<Declaration> declarations)
        {
            Member = member;

            MemberParams = declarations.Where(item => item.DeclarationType == DeclarationType.Parameter &&
                                          item.ParentScope == Member.Scope)
                                       .OrderBy(o => o.Selection.StartLine)
                                       .ThenBy(t => t.Selection.StartColumn)
                                       .Select(p => new Parameter
                                                    {
                                                        ParamAccessibility = p.Accessibility,
                                                        ParamName = p.IdentifierName,
                                                        ParamType = p.AsTypeName
                                                    })
                                       .ToList();

            IsSelected = false;
        }

        public override string ToString()
        {
            return Member.IdentifierName + "(" + string.Join(", ", MemberParams.Select(m => m.ParamType)) + ")";
        }
    }
}
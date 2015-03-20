using Rubberduck.Parsing;

namespace Rubberduck.VBA.Nodes
{
    public class ParameterNode : Node
    {
        public enum VBParameterType
        {
            ImplicitByRef,
            ByRef,
            ByVal
        }

        public ParameterNode(VBAParser.ArgContext context, string scope)
            : base(context, scope)
        {
        }

        private new VBAParser.ArgContext Context { get { return Context;} }

        public VBParameterType PassedBy
        {
            get
            {
                return Context.BYVAL() != null 
                    ? VBParameterType.ByVal
                    : Context.BYREF() != null
                        ? VBParameterType.ByRef 
                        : VBParameterType.ImplicitByRef;
            }
        }

        public string Name { get { return Context.ambiguousIdentifier().GetText(); } }
        public string TypeName 
        { 
            get 
            { 
                return Context.asTypeClause() == null || string.IsNullOrEmpty(Context.asTypeClause().GetText())
                    ? Tokens.Variant
                    : Context.asTypeClause().type().GetText(); 
            } 
        }

        public bool IsOptional
        {
            get
            {
                return Context.OPTIONAL() == null || string.IsNullOrEmpty(Context.OPTIONAL().GetText());
            }
        }
    }
}

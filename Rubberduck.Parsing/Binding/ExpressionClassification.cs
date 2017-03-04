namespace Rubberduck.Parsing.Binding
{
    public enum ExpressionClassification
    {
        ResolutionFailed,
        Value,
        Variable,
        Property,
        Function,
        Subroutine,
        Unbound,
        Project,
        ProceduralModule,
        Type
    }
}

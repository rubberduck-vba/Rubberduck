namespace Rubberduck.Parsing.VBA
{
    public interface IParsingStageService: ICOMReferenceSynchronizer, IBuiltInDeclarationLoader, IParseRunner, IDeclarationResolveRunner, IReferenceResolveRunner 
    {
    }
}

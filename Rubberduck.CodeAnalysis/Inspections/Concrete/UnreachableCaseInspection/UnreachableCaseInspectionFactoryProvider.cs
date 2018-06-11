
namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionFactoryProvider
    {
        IParseTreeValueFactory CreateIParseTreeValueFactory();
        IUnreachableCaseInspectorFactory CreateIUnreachableInspectorFactory();
    }

    public class UnreachableCaseInspectionFactoryProvider : IUnreachableCaseInspectionFactoryProvider
    {
        public IParseTreeValueFactory CreateIParseTreeValueFactory()
        {
            return new ParseTreeValueFactory();
        }

        public IUnreachableCaseInspectorFactory CreateIUnreachableInspectorFactory()
        {
            return new UnreachableCaseInspectorFactory();
        }
    }
}

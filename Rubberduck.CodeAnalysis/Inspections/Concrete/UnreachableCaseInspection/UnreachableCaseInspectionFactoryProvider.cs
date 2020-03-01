
namespace Rubberduck.Inspections.Concrete.UnreachableCaseInspection
{
    public interface IUnreachableCaseInspectionFactoryProvider
    {
        IParseTreeValueFactory CreateIParseTreeValueFactory();
        IUnreachableCaseInspectorFactory CreateIUnreachableInspectorFactory();
        IParseTreeValueVisitorFactory CreateParseTreeValueVisitorFactory();
    }

    public class UnreachableCaseInspectionFactoryProvider : IUnreachableCaseInspectionFactoryProvider
    {
        public IParseTreeValueFactory CreateIParseTreeValueFactory()
        {
            return new ParseTreeValueFactory();
        }

        public IUnreachableCaseInspectorFactory CreateIUnreachableInspectorFactory()
        {
            return new UnreachableCaseInspectorFactory(CreateIParseTreeValueFactory());
        }

        public IParseTreeValueVisitorFactory CreateParseTreeValueVisitorFactory()
        {
            return new ParseTreeValueVisitorFactory(CreateIParseTreeValueFactory());
        }
    }
}

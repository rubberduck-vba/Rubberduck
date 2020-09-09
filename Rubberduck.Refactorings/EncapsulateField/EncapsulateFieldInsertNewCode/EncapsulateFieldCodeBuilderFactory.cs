
namespace Rubberduck.Refactorings.EncapsulateField
{
    public interface IEncapsulateFieldCodeBuilderFactory
    {
        IEncapsulateFieldCodeBuilder Create();
    }

    public class EncapsulateFieldCodeBuilderFactory : IEncapsulateFieldCodeBuilderFactory
    {
        private readonly ICodeBuilder _codeBuilder;
        public EncapsulateFieldCodeBuilderFactory(ICodeBuilder codeBuilder)
        {
            _codeBuilder = codeBuilder;
        }

        public IEncapsulateFieldCodeBuilder Create()
        {
            return new EncapsulateFieldCodeBuilder(_codeBuilder);
        }
    }
}

namespace Rubberduck.Parsing.Binding
{
    public sealed class ArgumentListArgument
    {
        private readonly ArgumentListArgumentType _argumentType;

        public ArgumentListArgument(ArgumentListArgumentType argumentType)
        {
            _argumentType = argumentType;
        }

        public ArgumentListArgumentType ArgumentType
        {
            get
            {
                return _argumentType;
            }
        }
    }
}

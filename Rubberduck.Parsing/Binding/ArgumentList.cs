using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Binding
{
    public sealed class ArgumentList
    {
        private readonly List<ArgumentListArgument> _arguments;

        public ArgumentList()
        {
            _arguments = new List<ArgumentListArgument>();
        }

        public void AddArgument(IExpressionBinding binding, ArgumentListArgumentType argumentType)
        {
            _arguments.Add(new ArgumentListArgument(binding, argumentType));
        }

        public bool HasArguments
        {
            get
            {
                return HasRequiredPositionalArgument|| HasNamedArguments;
            }
        }

        public bool HasNamedArguments
        {
            get
            {
                return _arguments.Any(a => a.ArgumentType == ArgumentListArgumentType.Named);
            }
        }

        public bool HasRequiredPositionalArgument
        {
            get
            {
                return _arguments.Any(a => a.ArgumentType == ArgumentListArgumentType.Positional);
            }
        }

        public IReadOnlyList<ArgumentListArgument> Arguments
        {
            get
            {
                return _arguments;
            }
        }
    }
}

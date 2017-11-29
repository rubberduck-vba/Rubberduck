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

        public void AddArgument(ArgumentListArgument argument)
        {
            _arguments.Add(argument);
        }

        public bool HasArguments => HasRequiredPositionalArgument || HasNamedArguments;

        public bool HasNamedArguments => _arguments.Any(a => a.ArgumentType == ArgumentListArgumentType.Named);

        public bool HasRequiredPositionalArgument => _arguments.Any(a => a.ArgumentType == ArgumentListArgumentType.Positional);

        public IReadOnlyList<ArgumentListArgument> Arguments => _arguments;
    }
}

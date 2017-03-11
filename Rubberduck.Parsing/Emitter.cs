using System.Diagnostics;
using System.Runtime.InteropServices;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing
{
    /// <summary>
    /// A hack up our sleeves to get the type name of document modules. <see cref="ExecuteWithResult{TResult}"/>.
    /// </summary>
    public class Emitter
    {
        private readonly RubberduckParserState _state;

        public Emitter(RubberduckParserState state)
        {
            _state = state;
        }

        private static readonly string GetTypeNameFunctionTemplate = @"
Public Function GetTypeName() As String
    GetTypeName = TypeName({0})
End Function
";

        /// <summary>
        /// Gets a VBA function that returns the name of the type of an expression specified as an argument.
        /// </summary>
        /// <param name="arg">The literal contents of the 'TypeName' argument list, in VBA runtime context.</param>
        /// <remarks>The argument list excludes the surrounding parentheses that delimit it.</remarks>
        /// <returns></returns>
        public string GetTypeNameFunctionBody(string arg)
        {
            return string.Format(GetTypeNameFunctionTemplate, arg);
        }

        private static readonly object ThreadLock = new object();

        /// <summary>
        /// Emits specified code into a new, temporary modules, executes specified function, returns the result and destroys the temporary module.
        /// </summary>
        /// <typeparam name="TResult">The result type.</typeparam>
        /// <param name="project">The project to execute the code in.</param>
        /// <param name="content">The content of the module to emit.</param>
        /// <param name="name">The function to execute.</param>
        /// <param name="args">The arguments to be passed to the function.</param>
        /// <returns>The result of the function, or the default value for the specified return type.</returns>
        public TResult ExecuteWithResult<TResult>(IVBProject project, string content, string name, params object[] args)
        {
            lock (ThreadLock)
            {
                Debug.Assert(content.Contains(Tokens.Public + ' ' + Tokens.Function + ' ' + name));
                Debug.Assert(project.Protection == ProjectProtection.Unprotected);

                _state.IsEnabled = false;
                IVBComponent component = null;
                object result;
                try
                {
                    component = project.VBComponents.Add(ComponentType.StandardModule);
                    component.CodeModule.AddFromString(content);
                    var host = project.VBE.HostApplication();
                    result = host.Run(name, args);
                }
                catch (COMException)
                {
                    // IHostApplication.Run is supported, but the call failed.
                    return default(TResult);
                }
                finally
                {
                    if (component != null)
                    {
                        project.VBComponents.Remove(component);
                    }
                }

                _state.IsEnabled = true;
                if (result == null)
                {
                    return default(TResult);
                }
                return (TResult) result;
            }
        }
    }
}

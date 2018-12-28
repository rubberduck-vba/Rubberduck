using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Reflection;
using Castle.Core;
using Castle.Facilities.TypedFactory;
using Castle.MicroKernel;
using Castle.MicroKernel.Context;
using Moq;
using Rubberduck.Refactorings;

namespace RubberduckTests.Refactoring.MockIoC
{
    internal class AutoMoqResolver : ISubDependencyResolver
    {
        private readonly IKernel _kernel;

        public AutoMoqResolver(IKernel kernel)
        {
            _kernel = kernel;
        }

        public bool CanResolve(
            CreationContext context,
            ISubDependencyResolver contextHandlerResolver,
            ComponentModel model,
            DependencyModel dependency)
        {
            if (dependency.TargetType == null)
                return false;

            Debug.Assert(dependency.TargetType.Name != "IRefactoringDialog");
            Debug.Assert(dependency.TargetType.Name != "RefactoringDialogStub");
            Debug.Assert(dependency.TargetType.Name != "IRefactoringView");
            Debug.Assert(dependency.TargetType.Name != "RefactoringViewStub");
            Debug.Assert(dependency.TargetType.Name != "IRefactoringViewModel");

            if (dependency.TargetType.Namespace == null)
                return false;

            if (!dependency.TargetType.Namespace.StartsWith("Rubberduck"))
                return false;

            if (dependency.TargetType.Name.EndsWith("Factory"))
                return false;

            if(dependency.TargetType.Name.StartsWith("IRefactoringView"))
                return true;

            if (dependency.TargetType.Name.StartsWith("IRefactoringDialog"))
                return false;

            return dependency.TargetType.IsInterface;
        }

        public object Resolve(
            CreationContext context,
            ISubDependencyResolver contextHandlerResolver,
            ComponentModel model,
            DependencyModel dependency)
        {
            var mockType = typeof(Mock<>).MakeGenericType(dependency.TargetType);
            return ((Mock)_kernel.Resolve(mockType)).Object;
        }
    }

    internal class AutoMoqFactorySelector : DefaultTypedFactoryComponentSelector
    {
        protected override Func<IKernelInternal, IReleasePolicy, object> BuildFactoryComponent(MethodInfo method, string componentName, Type componentType,
            IDictionary additionalArguments)
        {
            if (!componentType.IsGenericType)
            {
                return base.BuildFactoryComponent(method, componentName, componentType, additionalArguments);
            }

            if(componentType.GetGenericTypeDefinition() == typeof(IRefactoringDialog<,,>))
            {
                return (kernel, rp) =>
                {
                    var modelType = componentType.GenericTypeArguments[0];
                    var viewType = componentType.GenericTypeArguments[1];
                    var viewModelType = componentType.GenericTypeArguments[2];
                    var stubType =
                        typeof(RefactoringDialogStub<,,>).MakeGenericType(modelType, viewType, viewModelType);
                    var mockType =
                        typeof(Mock<>).MakeGenericType(stubType);

                    var args = new object[additionalArguments.Count];
                    additionalArguments.Values.CopyTo(args, 0);
                    var mockArgs =
                        new Dictionary<string, object>
                        {
                            {"behavior", MockBehavior.Default},
                            {"args", args}
                        };

                    var mock = (Mock) kernel.Resolve(mockType, mockArgs);
                    mock.CallBase = true;

                    return mock.Object;
                };
            }

            if (componentType.GetGenericTypeDefinition() == typeof(IRefactoringView<>))
            {
                return (kernel, rp) =>
                {
                    var modelType = componentType.GenericTypeArguments[0];
                    var stubType = typeof(RefactoringViewStub<>).MakeGenericType(modelType);
                    var mockType = typeof(Mock<>).MakeGenericType(stubType);
                    var args = new object[additionalArguments.Count];
                    additionalArguments.Values.CopyTo(args, 0);
                    var mockArgs =
                        new Dictionary<string, object>
                        {
                            {"behavior", MockBehavior.Default },
                            {"args", args }
                        };
                    var mock = (Mock) kernel.Resolve(mockType, mockArgs);
                    mock.CallBase = true;

                    return mock.Object;
                };
            }
            return base.BuildFactoryComponent(method, componentName, componentType, additionalArguments);
        }
    }
}

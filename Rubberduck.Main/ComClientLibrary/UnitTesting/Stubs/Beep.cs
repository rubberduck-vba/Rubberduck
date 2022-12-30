using Rubberduck.Runtime;

namespace Rubberduck.UnitTesting
{
    internal class Beep : StubBase
    {
        private readonly IBeepInterceptor _interceptor;

        public Beep(IBeepInterceptor interceptor)
        {
            _interceptor = interceptor;
            _interceptor.Beep += BeepCallback;
        }

        public void BeepCallback(object sender, BeepEventArgs e)
        {
            OnCallBack(true);

            if (PassThrough)
            {
                _interceptor.NativeCall();
            }

            e.Handled = true;
        }

        public override void Dispose()
        {
            _interceptor.Beep -= BeepCallback;
            base.Dispose();
        }
    }
}

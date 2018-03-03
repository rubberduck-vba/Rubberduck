namespace Rubberduck.UnitTesting
{
    public interface IFakes
    {
        void StartTest();
        void StopTest();
    }

    public interface IFakesFactory
    {
        IFakes GetFakes();
    }
}

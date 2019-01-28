namespace Rubberduck.Deployment.Build.Structs
{
    public enum Bitness
    {
        /// <summary>Indicates that the bitness does not matter; the data will be same for all platforms. </summary>
        IsAgnostic,
        /// <summary>Indicates that different versions of data should be present for each platform's bitness. </summary>
        IsPlatformDependent,
        /// <summary>Indicates that it should be always present irrespective of the platform but contains 32-bit specific data. </summary>
        Is32Bit,
        /// <summary>Indicates that it should be always present irrespective of the platform but contains 32-bit specific data. Should be ignored on 32-bit OS. </summary>
        Is64Bit
    }
}
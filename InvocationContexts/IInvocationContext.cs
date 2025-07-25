namespace Printune
{
    /// <summary>
    /// A super simple interface to allow for interchangeability of invocation contexts.
    /// This is largely what enables the commandline application to run in different modes without a ton of spaghetti code.
    /// </summary>
    public interface IInvocationContext
    {
        /// <summary>
        /// The only requirement of the IInvocation interface.
        /// </summary>
        /// <returns>The integer that Program.cs->Main() returns upon exit.</returns>
        int Invoke();
    }
}
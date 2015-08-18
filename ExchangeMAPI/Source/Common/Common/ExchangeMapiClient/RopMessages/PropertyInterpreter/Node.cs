namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// Abstract class defines the interface Parse,
    /// which must be implemented by all sub classes 
    /// </summary>
    public abstract class Node
    {
        /// <summary>
        /// Parse bytes in context into a Node
        /// </summary>
        /// <param name="context">The Context</param>
        public abstract void Parse(Context context);
    }
}
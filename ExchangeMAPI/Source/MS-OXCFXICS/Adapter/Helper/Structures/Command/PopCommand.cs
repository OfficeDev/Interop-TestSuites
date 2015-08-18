namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// Represent a pop command.
    /// </summary>
    public class PopCommand : Command
    {
        /// <summary>
        /// Initializes a new instance of the PopCommand class.
        /// </summary>
        /// <param name="command">The command byte.</param>
        public PopCommand(byte command) :
            base(command, 0x50, 0x50)
        { 
        }
    }
}
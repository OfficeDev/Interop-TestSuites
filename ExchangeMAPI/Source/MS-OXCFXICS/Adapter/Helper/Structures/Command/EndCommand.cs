namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// Represent an end command.
    /// </summary>
    public class EndCommand : Command
    {
        /// <summary>
        /// Initializes a new instance of the EndCommand class.
        /// </summary>
        /// <param name="command">The command byte.</param>
        public EndCommand(byte command)
            : base(command, 0x00, 0x00)
        { 
        }
    }
}
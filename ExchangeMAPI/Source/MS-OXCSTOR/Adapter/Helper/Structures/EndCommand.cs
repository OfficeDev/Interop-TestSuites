namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    /// <summary>
    /// EndCommand class
    /// </summary>
    public class EndCommand : BaseCommand
    {
        /// <summary>
        /// Initializes a new instance of the EndCommand class
        /// </summary>
        public EndCommand()
        {
            this.Command = (byte)CommandType.EndCommand;
        }

        /// <summary>
        /// Get the size of the EndCommand
        /// </summary>
        /// <returns>The size of the EndCommand</returns>
        public override int Size()
        {
            return 1;
        }

        /// <summary>
        /// Get the bytes of the EndCommand
        /// </summary>
        /// <returns>The bytes of the EndCommand</returns>
        public override byte[] GetBytes()
        {
            byte[] resultBytes = new byte[1];
            resultBytes[0] = 0x00;
            return resultBytes;
        }
    }
}
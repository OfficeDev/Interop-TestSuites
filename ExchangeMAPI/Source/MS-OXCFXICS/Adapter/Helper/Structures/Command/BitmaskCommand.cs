namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    /// <summary>
    /// Represent a bitmask command.
    /// </summary>
    public class BitmaskCommand : Command
    {
        /// <summary>
        /// The low-order byte of the low value of the first GLOBCNT range.
        /// </summary>
        private byte startValue;

        /// <summary>
        /// One bit set for each value within a range, 
        /// excluding the low value of the first GLOBCNT range.
        /// </summary>
        private byte bitmask;

        /// <summary>
        /// Initializes a new instance of the BitmaskCommand class.
        /// </summary>
        /// <param name="command">The command byte.</param>
        /// <param name="startValue">Low-order byte of first GLOBCNT.</param>
        /// <param name="bitmask">Used for GLOBCNT generation where values are defined based 
        /// on the StartingValue and which bits are set in Bitmask.
        /// </param>
        public BitmaskCommand(byte command, byte startValue, byte bitmask)
            : base(command, 0x42, 0x42)
        {
            this.startValue = startValue;
            this.bitmask = bitmask;
        }

        /// <summary>
        /// Gets the stsrtValue.
        /// Low-order byte of first GLOBCNT.
        /// </summary>
        public byte StartValue
        {
            get { return this.startValue; }
        }

        /// <summary>
        /// Gets the bitmask.
        /// Used for GLOBCNT generation where values are defined based 
        /// on the StartingValue and which bits are set in Bitmask.
        /// </summary>
        public byte Bitmask
        {
            get { return this.bitmask; }
        }
    }
}
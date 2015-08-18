namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// Represent a range command.
    /// </summary>
    public class RangeCommand : Command
    {
        /// <summary>
        /// The low value of the range.
        /// </summary>
        private byte[] lowValue;

        /// <summary>
        /// The high value of the range.
        /// </summary>
        private byte[] highValue;

        /// <summary>
        /// Initializes a new instance of the RangeCommand class.
        /// </summary>
        /// <param name="command">The command byte.</param>
        /// <param name="lowValue">Variable length byte array of low-order values for GLOBCNT generation.</param>
        /// <param name="highValue">Variable length byte array of high-order values for GLOBCNT generation.</param>
        public RangeCommand(byte command, byte[] lowValue, byte[] highValue)
            : base(command, 0x52, 0x52)
        {
            AdapterHelper.Site.Assert.AreEqual(lowValue.Length, highValue.Length, "The lowValue length and highValue length are not equal, the lowValue length is {0} and highValue length is {1}.", lowValue.Length, highValue.Length);

            this.lowValue = lowValue;
            this.highValue = highValue;
        }

        /// <summary>
        /// Gets variable length byte array of low-order values for GLOBCNT generation.
        /// </summary>
        public byte[] LowValue
        {
            get { return this.lowValue; }
        }

        /// <summary>
        ///  Gets variable length byte array of low-order values for GLOBCNT generation.
        /// </summary>
        public byte[] HighValue
        {
            get { return this.highValue; }
        }
    }
}
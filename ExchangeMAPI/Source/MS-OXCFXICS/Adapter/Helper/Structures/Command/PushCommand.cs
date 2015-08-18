namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System;

    /// <summary>
    /// Represent a push command.
    /// </summary>
    public class PushCommand : Command
    {
        /// <summary>
        /// Common bytes pushed.
        /// </summary>
        private byte[] commonBytes;

        /// <summary>
        /// Initializes a new instance of the PushCommand class.
        /// </summary>
        /// <param name="command">The command byte.</param>
        /// <param name="commonBytes">The common bytes pushed by this command.</param>
        public PushCommand(byte command, byte[] commonBytes) :
            base(command, 1, 6)
        {
            if (!this.CheckCommand(command, 1, 6)
                || commonBytes.Length != command)
            {
                AdapterHelper.Site.Assert.Fail("The command is invalid.");
            }

            this.commonBytes = commonBytes;
        }

        /// <summary>
        /// Gets the common bytes pushed by this command.
        /// </summary>
        public byte[] CommonBytes
        {
            get { return this.commonBytes; }
        }
    }
}
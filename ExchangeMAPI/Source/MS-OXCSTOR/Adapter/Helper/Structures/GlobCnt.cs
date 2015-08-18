namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    /// <summary>
    /// Global Change
    /// </summary>
    public class GlobCnt
    {
        /// <summary>
        /// Command type
        /// </summary>
        private CommandType type;

        /// <summary>
        /// Base command
        /// </summary>
        private BaseCommand command;

        /// <summary>
        /// Gets or sets the type
        /// </summary>
        public CommandType Type
        {
            get
            {
                return this.type;
            }

            set
            {
                this.type = value;
            }
        }

        /// <summary>
        /// Gets or sets the command
        /// </summary>
        public BaseCommand Command
        {
            get
            {
                return this.command;
            }

            set
            {
                this.command = value;
            }
        }
    }
}
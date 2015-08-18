namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    using System;

    /// <summary>
    /// PushCommand class
    /// </summary>
    public class PushCommand : BaseCommand
    {
        /// <summary>
        /// Initializes a new instance of the PushCommand class
        /// </summary>
        public PushCommand()
        {
        }

        /// <summary>
        /// Initializes a new instance of the PushCommand class
        /// </summary>
        /// <param name="data">The data that will be pushed</param>
        public PushCommand(byte[] data)
        {
            this.CommandBytes = new byte[data.Length];
            Array.Copy(data, this.CommandBytes, data.Length);
        }

        /// <summary>
        /// Generate random bytes in command
        /// </summary>
        public void GenerateRandomCommandBytes()
        {
            this.CommandBytes = new byte[Command];
            Random rnd = new Random((int)DateTime.Now.Ticks);
            rnd.NextBytes(this.CommandBytes);
        }

        /// <summary>
        /// Get the size of the PushCommand
        /// </summary>
        /// <returns>The size of the PushCommand</returns>
        public override int Size()
        {
            return 1 + this.CommandBytes.Length;
        }

        /// <summary>
        /// Get the bytes of the PushCommand
        /// </summary>
        /// <returns>The bytes of the PushCommand</returns>
        public override byte[] GetBytes()
        {
            byte[] resultBytes = new byte[1 + this.CommandBytes.Length];
            resultBytes[0] = this.Command;
            Array.Copy(this.CommandBytes, 0, resultBytes, 1, this.CommandBytes.Length);
            return resultBytes;
        }
    }
}
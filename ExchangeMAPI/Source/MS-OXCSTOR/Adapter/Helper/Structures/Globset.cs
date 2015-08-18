namespace Microsoft.Protocols.TestSuites.MS_OXCSTOR
{
    using System;
    using System.Collections.Generic;

    /// <summary>
    /// A GLOBSET is a set of GLOBCNT values that are typically reduced to GLOBCNT ranges.
    /// </summary>
    public class Globset
    {
        /// <summary>
        /// A set of GLOBCNT values
        /// </summary>
        private List<GlobCnt> globCntList = new List<GlobCnt>();

        /// <summary>
        /// Initializes a new instance of the Globset class
        /// </summary>
        public Globset()
        {
        }

        /// <summary>
        /// Initializes a new instance of the Globset class
        /// </summary>
        /// <param name="data">Byte array value</param>
        public Globset(byte[] data)
        {
            int index = 0;

            while (index <= data.Length - 1)
            {
                int commnadSize = this.ParseCommand(data, index);
                index += commnadSize;
            }
        }

        /// <summary>
        /// Gets the count of Globcnt
        /// </summary>
        public int SetCount
        {
            get
            {
                return this.globCntList.Count;
            }
        }

        /// <summary>
        /// Sets the globCntList
        /// </summary>
        public List<GlobCnt> GlobCntList
        {
            set
            {
                this.globCntList = value;
            }
        }

        /// <summary>
        /// Get Globset in bytes 
        /// </summary>
        /// <returns>Globset in bytes </returns>
        public byte[] GetBytes()
        {
            int size = 0;

            for (int i = 0; i < this.globCntList.Count; i++)
            {
                size += this.globCntList[i].Command.Size();
            }

            byte[] resultBytes = new byte[size];

            int index = 0;
            for (int i = 0; i < this.globCntList.Count; i++)
            {
                Array.Copy(this.globCntList[i].Command.GetBytes(), 0, resultBytes, index, this.globCntList[i].Command.Size());
                index += this.globCntList[i].Command.Size();
            }

            return resultBytes;
        }

        /// <summary>
        /// Parse a command by given the data
        /// </summary>
        /// <param name="data">Byte array value</param>
        /// <param name="index">The starting index</param>
        /// <returns>The size of the data structure</returns>
        private int ParseCommand(byte[] data, int index)
        {
            int cmdSize = 0;
            byte type = data[index++];
            cmdSize++;

            // Push command
            if (type <= (byte)CommandType.PushCommand && type != (byte)CommandType.EndCommand)
            {
                GlobCnt globc = new GlobCnt
                {
                    Type = CommandType.PushCommand
                };
                byte[] cmdData = new byte[type];
                Array.Copy(data, index, cmdData, 0, (int)type);
                index += (int)type;
                cmdSize += (int)type;
                globc.Command = new PushCommand(cmdData);
                this.globCntList.Add(globc);
            }
            else if (type == (byte)CommandType.PopCommand)
            {
                GlobCnt globc = new GlobCnt
                {
                    Type = CommandType.PopCommand, Command = new PopCommand()
                };
                this.globCntList.Add(globc);
                cmdSize = globc.Command.Size();
            }
            else if (type == (byte)CommandType.RangCommand)
            {
                GlobCnt globc = new GlobCnt
                {
                    Type = CommandType.RangCommand, Command = new RangeCommand()
                };
                this.globCntList.Add(globc);
                cmdSize = globc.Command.Size();
            }
            else if (type == (byte)CommandType.EndCommand)
            {
                // The End Command must occurs at the last
                if (index != data.Length)
                {
                    throw new Exception("Not an valid IDSET");
                }

                GlobCnt globc = new GlobCnt
                {
                    Type = CommandType.EndCommand, Command = new EndCommand()
                };
                this.globCntList.Add(globc);
            }
            else
            {
                throw new Exception("Currently, only support pushCommand");
            }

            return cmdSize;
        }
    }
}
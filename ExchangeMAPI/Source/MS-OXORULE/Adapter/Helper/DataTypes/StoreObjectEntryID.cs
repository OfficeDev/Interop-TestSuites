namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// A Store object EntryID identifies a mailbox Store object or a public folder Store object itself, rather than a message or Folder object residing in such a database. 
    /// </summary>
    public class StoreObjectEntryID
    {
        /// <summary>
        /// The flags that MUST be 0x00000000.
        /// </summary>
        public readonly uint Flags = 0x00000000;

        /// <summary>
        /// The ProviderUID that MUST be %x38.A1.BB.10.05.E5.10.1A.A1.BB.08.00.2B.2A.56.C2.
        /// </summary>
        public readonly byte[] ProviderUID = new byte[] { 0x38, 0xA1, 0xBB, 0x10, 0x05, 0xE5, 0x10, 0x1A, 0xA1, 0xBB, 0x08, 0x00, 0x2B, 0x2A, 0x56, 0xC2 };

        /// <summary>
        /// MUST be zero.
        /// </summary>
        public readonly byte Version = 0x00;

        /// <summary>
        /// The flag value that MUST be zero.
        /// </summary>
        public readonly byte Flag = 0x00;

        /// <summary>
        /// MUST be set to the following value which represents "emsmdb.dll": %x45.4D.53.4D.44.42.2E.44.4C.4C.00.00.00.00.
        /// </summary>
        public readonly byte[] DLLFileName = new byte[] { 0x45, 0x4D, 0x53, 0x4D, 0x44, 0x42, 0x2E, 0x44, 0x4C, 0x4C, 0x00, 0x00, 0x00, 0x00 };

        /// <summary>
        /// MUST be 0x00000000.
        /// </summary>
        public readonly uint WrappedFlags = 0x00000000;

        /// <summary>
        /// WrappedProviderUID value of Mailbox.
        /// </summary>
        private readonly byte[] mailboxWrappedProviderUID = new byte[] { 0x1B, 0x55, 0xFA, 0x20, 0xAA, 0x66, 0x11, 0xCD, 0x9B, 0xC8, 0x00, 0xAA, 0x00, 0x2F, 0xC4, 0x5A };

        /// <summary>
        /// WrappedProviderUID value of public folder.
        /// </summary>
        private readonly byte[] publicFolderWrappedProviderUID = new byte[] { 0x1C, 0x83, 0x02, 0x10, 0xAA, 0x66, 0x11, 0xCD, 0x9B, 0xC8, 0x00, 0xAA, 0x00, 0x2F, 0xC4, 0x5A };

        /// <summary>
        /// Type of Store object
        /// </summary>
        private StoreObjectType objectType;

        /// <summary>
        /// A string of single-byte characters terminated by a single zero byte, indicating the shortname or NetBIOS name of the server.
        /// </summary>
        private string serverShortname;

        /// <summary>
        /// A string of single-byte characters terminated by a single zero byte and representing the X500 DN of the mailbox, as specified in [MS-OXOAB]. This field is present only for mailbox databases.
        /// </summary>
        private string mailBoxDN;

        /// <summary>
        /// Initializes a new instance of the StoreObjectEntryID class.
        /// </summary>
        /// <param name="objType">Type of Store object.</param>
        public StoreObjectEntryID(StoreObjectType objType)
        {
            this.objectType = objType;
        }

        /// <summary>
        /// Initializes a new instance of the StoreObjectEntryID class.
        /// </summary>
        public StoreObjectEntryID()
        {
            this.objectType = StoreObjectType.Mailbox;
        }

        /// <summary>
        /// Gets type of Store object
        /// </summary>
        public StoreObjectType ObjectType
        {
            get
            {
                return this.objectType;
            }
        }

        /// <summary>
        /// Gets or sets a string of single-byte characters terminated by a single zero byte, indicating the shortname or NetBIOS name of the server.
        /// </summary>
        public string ServerShortname
        {
            get { return this.serverShortname; }
            set { this.serverShortname = value; }
        }

        /// <summary>
        /// Gets or sets a string of single-byte characters terminated by a single zero byte and representing the X500 DN of the mailbox, as specified in [MS-OXOAB]. This field is present only for mailbox databases.
        /// </summary>
        public string MailBoxDN
        {
            get { return this.mailBoxDN; }
            set { this.mailBoxDN = value; }
        }

        /// <summary>
        /// Gets the WrappedProviderUID that MUST be one of the following values:
        /// mailbox store object:%x1B.55.FA.20.AA.66.11.CD.9B.C8.00.AA.00.2F.C4.5A
        /// public folder store object:%x1C.83.02.10.AA.66.11.CD.9B.C8.00.AA.00.2F.C4.5A
        /// </summary>
        public byte[] WrappedProviderUID
        {
            get
            {
                if (this.ObjectType == StoreObjectType.Mailbox)
                {
                    return this.mailboxWrappedProviderUID;
                }
                else
                {
                    return this.publicFolderWrappedProviderUID;
                }
            }
        }

        /// <summary>
        /// Gets WrappedType that MUST be %x0C.00.00.00 for a mailbox store, or %x06.00.00.00 for a public store.
        /// </summary>
        public uint WrappedType
        {
            get
            {
                if (this.ObjectType == StoreObjectType.Mailbox)
                {
                    return 0x0000000C;
                }
                else
                {
                    return 0x00000006;
                }
            }
        }

        /// <summary>
        /// Get size of this class
        /// </summary>
        /// <returns>Return the size value.</returns>
        public int Size()
        {
            return this.Serialize().Length;
        }

        /// <summary>
        /// Get serialized byte array for this struct
        /// </summary>
        /// <returns>Return the result of serialize.</returns>
        public byte[] Serialize()
        {
            List<byte> bytes = new List<byte>();
            bytes.AddRange(BitConverter.GetBytes(this.Flags));
            bytes.AddRange(this.ProviderUID);
            bytes.Add(this.Version);
            bytes.Add(this.Flag);
            bytes.AddRange(this.DLLFileName);
            bytes.AddRange(BitConverter.GetBytes(this.WrappedFlags));
            bytes.AddRange(this.WrappedProviderUID);
            bytes.AddRange(BitConverter.GetBytes(this.WrappedType));
            bytes.AddRange(Encoding.ASCII.GetBytes(this.ServerShortname + "\0"));
            bytes.AddRange(Encoding.ASCII.GetBytes(this.MailBoxDN + "\0"));
            return bytes.ToArray();
        }

        /// <summary>
        /// Deserialized byte array to an ActionBlock instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of an ActionBlock instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public uint Deserialize(byte[] buffer)
        {
            BufferReader reader = new BufferReader(buffer);
            uint flags = reader.ReadUInt32();
            if (this.Flags != flags)
            {
                string errorMessage = "Flags MUST be 0x00000000, the actual value is " + flags.ToString() + "!";
                throw new ArgumentException(errorMessage);
            }

            byte[] providerUID = reader.ReadBytes(16);
            if (!Common.CompareByteArray(this.ProviderUID, providerUID))
            {
                string errorMessage = "ProviderUID MUST be %x38.A1.BB.10.05.E5.10.1A.A1.BB.08.00.2B.2A.56.C2., the actual value is " + providerUID.ToString() + "!";
                throw new ArgumentException(errorMessage);
            }

            byte version = reader.ReadByte();
            if (this.Version != version)
            {
                string errorMessage = "Version MUST be zero., the actual value is " + version.ToString() + "!";
                throw new ArgumentException(errorMessage);
            }

            byte flag = reader.ReadByte();
            if (this.Flag != flag)
            {
                string errorMessage = "Flag MUST be zero, the actual value is " + flag.ToString() + "!";
                throw new ArgumentException(errorMessage);
            }

            byte[] dllFileName = reader.ReadBytes(14);
            if (!Common.CompareByteArray(this.DLLFileName, dllFileName))
            {
                string errorMessage = "DLLFileName MUST be set to the following value which represents \"emsmdb.dll\": %x45.4D.53.4D.44.42.2E.44.4C.4C.00.00.00.00., the actual value is " + dllFileName.ToString() + "!";
                throw new ArgumentException(errorMessage);
            }

            uint wrappedFlags = reader.ReadUInt32();
            if (this.WrappedFlags != wrappedFlags)
            {
                string errorMessage = "WrappedFlags MUST be 0x00000000, the actual value is " + wrappedFlags.ToString() + "!";
                throw new ArgumentException(errorMessage);
            }

            byte[] wrappedProviderUID = reader.ReadBytes(16);
            if (Common.CompareByteArray(this.mailboxWrappedProviderUID, wrappedProviderUID))
            {
                this.objectType = StoreObjectType.Mailbox;
            }
            else if (Common.CompareByteArray(this.publicFolderWrappedProviderUID, wrappedProviderUID))
            {
                this.objectType = StoreObjectType.PublicFolder;
            }
            else
            {
                string errorMessage = "WrappedProviderUID is not mailbox or public folder, the actual wrappedProviderUID value is " + wrappedProviderUID.ToString() + "!";
                throw new ArgumentException(errorMessage);
            }

            uint wrappedType = reader.ReadUInt32();
            if (this.WrappedType != wrappedType)
            {
                if (this.ObjectType == StoreObjectType.Mailbox)
                {
                    string errorMessage = "For Mailbox Store object, WrappedType MUST be %x0C.00.00.00, the actual value is " + StoreObjectType.Mailbox.ToString() + "!";
                    throw new ArgumentException(errorMessage);
                }
                else
                {
                    string errorMessage = "For Public folder Store object, WrappedType MUST be %x06.00.00.00, the actual value is " + wrappedType.ToString() + "!";
                    throw new ArgumentException(errorMessage);
                }
            }

            this.ServerShortname = reader.ReadASCIIString();
            this.MailBoxDN = reader.ReadASCIIString();
            return reader.Position;
        }
    }
}
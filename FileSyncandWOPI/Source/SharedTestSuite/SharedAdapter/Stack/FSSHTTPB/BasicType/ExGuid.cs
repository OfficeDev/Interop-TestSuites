namespace Microsoft.Protocols.TestSuites.SharedAdapter
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The class is used to specify the ExGuid structure.
    /// </summary>
    public class ExGuid : BasicObject
    {
        /// <summary>
        /// Specify the extended GUID null type value.
        /// </summary>
        public const int ExtendedGUIDNullType = 0;
        
        /// <summary>
        /// Specify the extended GUID 5 Bit uint type value.
        /// </summary>
        public const int ExtendedGUID5BitUintType = 4;
        
        /// <summary>
        /// Specify the extended GUID 10 Bit uint type value.
        /// </summary>
        public const int ExtendedGUID10BitUintType = 32;
        
        /// <summary>
        /// Specify the extended GUID 17 Bit uint type value.
        /// </summary>
        public const int ExtendedGUID17BitUintType = 64;
        
        /// <summary>
        /// Specify the extended GUID 32 Bit uint type value.
        /// </summary>
        public const int ExtendedGUID32BitUintType = 128;

        /// <summary>
        /// Initializes a new instance of the ExGuid class with specified value.
        /// </summary>
        /// <param name="value">Specify the ExGUID Value.</param>
        /// <param name="identifier">Specify the ExGUID GUID value.</param>
        public ExGuid(uint value, Guid identifier)
        {
            this.Value = value;
            this.GUID = identifier;
        }

        /// <summary>
        /// Initializes a new instance of the ExGuid class, this is the copy constructor.
        /// </summary>
        /// <param name="guid2">Specify the ExGuid instance where copies from.</param>
        public ExGuid(ExGuid guid2)
        {
            this.Value = guid2.Value;
            this.GUID = guid2.GUID;
            this.Type = guid2.Type;
        }

        /// <summary>
        /// Initializes a new instance of the ExGuid class, this is a default constructor.
        /// </summary>
        public ExGuid()
        {
            this.GUID = Guid.Empty;
        }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the type.
        /// </summary>
        public uint Type { get; set; }

        /// <summary>
        /// Gets or sets an unsigned integer that specifies the value.
        /// </summary>
        public uint Value { get; set; }

        /// <summary>
        /// Gets or sets a GUID that specifies the item. MUST NOT be "{00000000-0000-0000-0000-000000000000}".
        /// </summary>
        public Guid GUID { get; set; }

        /// <summary>
        /// This method is used to convert the element of ExGuid basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ExGuid.</returns>
        public override List<byte> SerializeToByteList()
        {
            BitWriter bitWriter = new BitWriter(21);

            if (this.GUID == Guid.Empty)
            {
                bitWriter.AppendUInit32(0, 8);
            }
            else if (this.Value >= 0x00 && this.Value <= 0x1F)
            {
                bitWriter.AppendUInit32(ExtendedGUID5BitUintType, 3);
                bitWriter.AppendUInit32(this.Value, 5);
                bitWriter.AppendGUID(this.GUID);
            }
            else if (this.Value >= 0x20 && this.Value <= 0x3FF)
            {
                bitWriter.AppendUInit32(ExtendedGUID10BitUintType, 6);
                bitWriter.AppendUInit32(this.Value, 10);
                bitWriter.AppendGUID(this.GUID);
            }
            else if (this.Value >= 0x400 && this.Value <= 0x1FFFF)
            {
                bitWriter.AppendUInit32(ExtendedGUID17BitUintType, 7);
                bitWriter.AppendUInit32(this.Value, 17);
                bitWriter.AppendGUID(this.GUID);
            }
            else if (this.Value >= 0x20000 && this.Value <= 0xFFFFFFFF)
            {
                bitWriter.AppendUInit32(ExtendedGUID32BitUintType, 8);
                bitWriter.AppendUInit32(this.Value, 32);
                bitWriter.AppendGUID(this.GUID);
            }

            return new List<byte>(bitWriter.Bytes);
        }

        /// <summary>
        /// Override the Equals method.
        /// </summary>
        /// <param name="obj">Specify the object.</param>
        /// <returns>Return true if equals, otherwise return false.</returns>
        public override bool Equals(object obj)
        {
            ExGuid another = obj as ExGuid;

            if (another == null)
            {
                return false;
            }

            if (this.GUID != null && another.GUID != null)
            {
                return another.GUID.Equals(this.GUID) && another.Value == this.Value;
            }

            return false;
        }

        /// <summary>
        /// Override the GetHashCode.
        /// </summary>
        /// <returns>Return the hash value.</returns>
        public override int GetHashCode()
        {
            return this.GUID.GetHashCode() + this.Value.GetHashCode();
        }

        /// <summary>
        /// This method is used to deserialize the ExGuid basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ExGuid basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex)
        {
            using (BitReader bitReader = new BitReader(byteArray, startIndex))
            {
                int numberOfContinousZeroBit = 0;
                while (numberOfContinousZeroBit < 8 && bitReader.MoveNext())
                {
                    if (bitReader.Current == false)
                    {
                        numberOfContinousZeroBit++;
                    }
                    else
                    {
                        break;
                    }
                }

                switch (numberOfContinousZeroBit)
                {
                    case 2:
                        this.Value = bitReader.ReadUInt32(5);
                        this.GUID = bitReader.ReadGuid();
                        this.Type = ExtendedGUID5BitUintType;
                        return 17;

                    case 5:
                        this.Value = bitReader.ReadUInt32(10);
                        this.GUID = bitReader.ReadGuid();
                        this.Type = ExtendedGUID10BitUintType;
                        return 18;

                    case 6:
                        this.Value = bitReader.ReadUInt32(17);
                        this.GUID = bitReader.ReadGuid();
                        this.Type = ExtendedGUID17BitUintType;
                        return 19;

                    case 7:
                        this.Value = bitReader.ReadUInt32(32);
                        this.GUID = bitReader.ReadGuid();
                        this.Type = ExtendedGUID32BitUintType;
                        return 21;

                    case 8:
                        this.GUID = Guid.Empty;
                        this.Type = ExtendedGUIDNullType;
                        return 1;

                    default:
                        throw new InvalidOperationException("Failed to parse the ExGuid, the type value is unexpected");
                }
            }
        }
    }

    /// <summary>
    /// This class is used to specify the Extended GUID Array.
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.StyleCop.CSharp.MaintainabilityRules", "SA1402:FileMayOnlyContainASingleClass", Justification = "Easy to maintain one group of classes in one .cs file.")]
    public class ExGUIDArray : BasicObject
    {
        /// <summary>
        /// Gets or sets an extended GUID array that specifies an array of items.
        /// </summary>
        private List<ExGuid> content = null;

        /// <summary>
        /// Initializes a new instance of the ExGUIDArray class with specified value.
        /// </summary>
        /// <param name="content">Specify the list of ExGuid contents.</param>
        public ExGUIDArray(List<ExGuid> content)
            : this()
        {
            this.content = new List<ExGuid>();
            if (content != null)
            {
                foreach (ExGuid extendGuid in content)
                {
                    this.content.Add(new ExGuid(extendGuid));
                }
            }

            this.Count.DecodedValue = (ulong)this.Content.Count;
        }

        /// <summary>
        /// Initializes a new instance of the ExGUIDArray class, this is copy constructor.
        /// </summary>
        /// <param name="extendGuidArray">Specify the ExGUIDArray where copies from.</param>
        public ExGUIDArray(ExGUIDArray extendGuidArray)
            : this(extendGuidArray.Content)
        {
        }

        /// <summary>
        /// Initializes a new instance of the ExGUIDArray class, this is the default constructor.
        /// </summary>
        public ExGUIDArray()
        {
            this.Count = new Compact64bitInt();
            this.content = new List<ExGuid>();
        }

        /// <summary>
        /// Gets or sets a compact unsigned 64-bit integer that specifies the count extended GUIDs in the array. 
        /// </summary>
        public Compact64bitInt Count { get; set; }

        /// <summary>
        /// Gets or sets an extended GUID array
        /// </summary>
        public List<ExGuid> Content
        {
            get
            {
                return this.content;
            }

            set
            {
                this.content = value;
                this.Count.DecodedValue = (ulong)this.Content.Count;
            }
        }

        /// <summary>
        /// This method is used to convert the element of ExGUIDArray basic object into a byte List.
        /// </summary>
        /// <returns>Return the byte list which store the byte information of ExGUIDArray.</returns>
        public override List<byte> SerializeToByteList()
        {
            this.Count.DecodedValue = (uint)this.content.Count;

            List<byte> result = new List<byte>();
            result.AddRange(this.Count.SerializeToByteList());
            foreach (ExGuid extendGuid in this.content)
            {
                result.AddRange(extendGuid.SerializeToByteList());
            }

            return result;
        }

        /// <summary>
        /// This method is used to deserialize the ExGUIDArray basic object from the specified byte array and start index.
        /// </summary>
        /// <param name="byteArray">Specify the byte array.</param>
        /// <param name="startIndex">Specify the start index from the byte array.</param>
        /// <returns>Return the length in byte of the ExGUIDArray basic object.</returns>
        protected override int DoDeserializeFromByteArray(byte[] byteArray, int startIndex) // return the length consumed
        {
            int index = startIndex;
            this.Count = BasicObject.Parse<Compact64bitInt>(byteArray, ref index);

            this.Content.Clear();
            for (uint i = 0; i < this.Count.DecodedValue; i++)
            {
                ExGuid temp = BasicObject.Parse<ExGuid>(byteArray, ref index);
                this.Content.Add(temp);
            }

            return index - startIndex;
        }
    }
}
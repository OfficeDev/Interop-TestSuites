namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Clients can use a CommentRestriction structure to save associated comments together with a restriction they pertain to. 
    /// The comments are formatted as an arbitrary array of TaggedPropValue structures, and servers MUST store and retrieve this information for the client. 
    /// If the Restriction field is present, servers MUST evaluate it; if it is not present, then the CommentRestriction node will effectively evaluate as TRUE. 
    /// In either case, the comments themselves have no effect on the evaluation of the restriction.
    /// </summary>
    public class CommentRestriction : Restrictions
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value specifies how many TaggedValue structures are present in TaggedValues.
        /// </summary>
        private byte taggedValuesCount;

        /// <summary>
        /// Array of TaggedPropertyValue structures. This field MUST contain TaggedValuesCount structures. The TaggedPropertyValue structures MUST NOT include any multi-valued properties.
        /// </summary>
        private TaggedPropertyValue[] taggedValues;

        /// <summary>
        /// Unsigned 8-bit integer. This field MUST contain either TRUE (0x01) or FALSE (0x00). A TRUE value means that the Restriction field is present, while a FALSE value indicates the Restriction field is not present.
        /// </summary>
        private byte restrictionPresent;

        /// <summary>
        /// (optional) A Restriction structure. This field is only present if RestrictionPresent is TRUE.
        /// </summary>
        private IRestriction restriction;

        /// <summary>
        /// Initializes a new instance of the CommentRestriction class.
        /// </summary>
        public CommentRestriction()
        {
            this.RestrictType = RestrictionType.CommentRestriction;
            this.CountType = CountByte.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the CommentRestriction class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class.</param>
        public CommentRestriction(CountByte countType)
        {
            this.RestrictType = RestrictionType.CommentRestriction;
            this.CountType = countType;
        }

        /// <summary>
        /// Gets or sets unsigned 8-bit integer. This value specifies how many TaggedValue structures are present in TaggedValues.
        /// </summary>
        public byte TaggedValuesCount
        {
            get { return this.taggedValuesCount; }
            set { this.taggedValuesCount = value; }
        }

        /// <summary>
        /// Gets or sets array of TaggedPropertyValue structures. This field MUST contain TaggedValuesCount structures. The TaggedPropertyValue structures MUST NOT include any multi-valued properties.
        /// </summary>
        public TaggedPropertyValue[] TaggedValues
        {
            get { return this.taggedValues; }
            set { this.taggedValues = value; }
        }

        /// <summary>
        /// Gets or sets unsigned 8-bit integer. This field MUST contain either TRUE (0x01) or FALSE (0x00). A TRUE value means that the Restriction field is present, while a FALSE value indicates the Restriction field is not present.
        /// </summary>
        public byte RestrictionPresent
        {
            get { return this.restrictionPresent; }
            set { this.restrictionPresent = value; }
        }

        /// <summary>
        /// Gets or sets a Restriction structure. This field is only present if RestrictionPresent is TRUE.
        /// </summary>
        public IRestriction Restriction
        {
            get { return this.restriction; }
            set { this.restriction = value; }
        }

        /// <summary>
        /// Get the total Size of Restriction
        /// </summary>
        /// <returns>The Size of Restriction buffer.</returns>
        public override int Size()
        {
            return this.Serialize().Length;
        }

        /// <summary>
        /// Get serialized byte array for this structure
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public override byte[] Serialize()
        {
            List<byte> bytes = new List<byte>
            {
                (byte)RestrictType, this.TaggedValuesCount
            };
            for (int i = 0; i < this.TaggedValuesCount; i++)
            {
                bytes.AddRange(this.TaggedValues[i].Serialize());
            }

            bytes.Add(this.RestrictionPresent);
            if (this.RestrictionPresent == 0x01)
            {
                bytes.AddRange(this.Restriction.Serialize());
            }

            return bytes.ToArray();
        }

        /// <summary>
        /// Deserialized byte array to a Restriction instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of a Restriction instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public override uint Deserialize(byte[] buffer)
        {
            BufferReader reader = new BufferReader(buffer);
            this.RestrictType = (RestrictionType)reader.ReadByte();
            this.TaggedValuesCount = reader.ReadByte();
            this.TaggedValues = new TaggedPropertyValue[this.TaggedValuesCount];
            uint size = reader.Position;
            byte[] tmpArray = reader.ReadToEnd();
            for (int i = 0; i < this.TaggedValuesCount; i++)
            {
                this.TaggedValues[i] = AdapterHelper.ReadTaggedProperty(tmpArray);
                uint tagLength = (uint)this.TaggedValues[i].Size();
                size += tagLength;
                reader = new BufferReader(tmpArray);
                tmpArray = reader.ReadBytes(tagLength, (uint)(tmpArray.Length - tagLength));
            }

            reader = new BufferReader(buffer);
            reader.ReadBytes(size);
            this.RestrictionPresent = reader.ReadByte();
            size += reader.Position;
            tmpArray = reader.ReadToEnd();
            if (this.RestrictionPresent == 0x01)
            {
                RestrictionType restrictionType = (RestrictionType)tmpArray[0];
                switch (restrictionType)
                {
                    case RestrictionType.AndRestriction:
                        this.Restriction = new AndRestriction(CountType);
                        break;
                    case RestrictionType.BitMaskRestriction:
                        this.Restriction = new BitMaskRestriction();
                        break;
                    case RestrictionType.CommentRestriction:
                        this.Restriction = new CommentRestriction(CountType);
                        break;
                    case RestrictionType.ComparePropertiesRestriction:
                        this.Restriction = new ComparePropertiesRestriction();
                        break;
                    case RestrictionType.ContentRestriction:
                        this.Restriction = new ContentRestriction();
                        break;
                    case RestrictionType.CountRestriction:
                        this.Restriction = new CountRestriction(CountType);
                        break;
                    case RestrictionType.ExistRestriction:
                        this.Restriction = new ExistRestriction();
                        break;
                    case RestrictionType.NotRestriction:
                        this.Restriction = new NotRestriction(CountType);
                        break;
                    case RestrictionType.OrRestriction:
                        this.Restriction = new OrRestriction(CountType);
                        break;
                    case RestrictionType.PropertyRestriction:
                        this.Restriction = new PropertyRestriction();
                        break;
                    case RestrictionType.SizeRestriction:
                        this.Restriction = new SizeRestriction();
                        break;
                    case RestrictionType.SubObjectRestriction:
                        this.Restriction = new SubObjectRestriction(CountType);
                        break;
                }

                size += this.Restriction.Deserialize(tmpArray);
            }

            return size;
        }
    }
}
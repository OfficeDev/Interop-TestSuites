namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The SubObjectRestriction structure applies its subrestriction to a Message object's attachment table or recipients. If ANY row of the subobject satisfies the subrestriction, then the message satisfies the SubObjectRestriction.
    /// </summary>
    public class SubObjectRestriction : Restrictions
    {
        /// <summary>
        /// This value is a PropertyTag that designates the target of the subrestriction Restriction. 
        /// </summary>
        private SubObjectValue subObject;

        /// <summary>
        /// A Restriction structure. This subrestriction is applied to the rows of the subobject.
        /// </summary>
        private IRestriction restriction;

        /// <summary>
        /// Initializes a new instance of the SubObjectRestriction class.
        /// </summary>
        public SubObjectRestriction()
        {
            this.RestrictType = RestrictionType.SubObjectRestriction;
            this.CountType = CountByte.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the SubObjectRestriction class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class.</param>
        public SubObjectRestriction(CountByte countType)
        {
            this.RestrictType = RestrictionType.SubObjectRestriction;
            this.CountType = countType;
        }

        /// <summary>
        /// Value of SubObject
        /// </summary>
        public enum SubObjectValue : uint
        {
            /// <summary>
            /// PropertyTag of PidTagMessageRecipients, this tag identifies all recipients of the current message.
            /// </summary>
            PidTagMessageRecipients = 0x0E12000D,

            /// <summary>
            /// PropertyTag of PidTagMessageAttachments, this tag identifies all attachments to the current message.
            /// </summary>
            PidTagMessageAttachments = 0x0E13000D
        }

        /// <summary>
        /// Gets or sets PropertyTag that designates the target of the subrestriction Restriction. 
        /// </summary>
        public SubObjectValue SubObject
        {
            get { return this.subObject; }
            set { this.subObject = value; }
        }

        /// <summary>
        /// Gets or sets a Restriction structure. This subrestriction is applied to the rows of the subobject.
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
                (byte)RestrictType, (byte)this.SubObject
            };
            bytes.AddRange(this.Restriction.Serialize());

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
            this.SubObject = (SubObjectValue)reader.ReadUInt32();

            uint size = reader.Position;
            byte[] tmpArray = reader.ReadToEnd();

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

            return size;
        }
    }
}
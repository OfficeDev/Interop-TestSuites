namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// The NotRestriction structure describes a NOT restriction, which is used to apply a logical NOT operation to a single restriction.
    /// The result of a NotRestriction is TRUE if the child restriction evaluates to FALSE, and FALSE if the child restriction evaluates to TRUE.
    /// </summary>
    public class NotRestriction : Restrictions
    {
        /// <summary>
        /// A restriction structure. This value specifies the restriction the logical NOT applies to.
        /// </summary>
        private IRestriction restriction;

        /// <summary>
        /// Initializes a new instance of the NotRestriction class.
        /// </summary>
        public NotRestriction()
        {
            this.RestrictType = RestrictionType.NotRestriction;
            this.CountType = CountByte.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the NotRestriction class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class.</param>
        public NotRestriction(CountByte countType)
        {
            this.RestrictType = RestrictionType.NotRestriction;
            this.CountType = countType;
        }

        /// <summary>
        /// Gets or sets a restriction structure. This value specifies the restriction the logical NOT applies to.
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
                (byte)this.RestrictType
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
            BufferReader bufferReader = new BufferReader(buffer);
            this.RestrictType = (RestrictionType)bufferReader.ReadByte();

            uint size = bufferReader.Position;
            byte[] tmpArray = bufferReader.ReadToEnd();

            RestrictionType restrictionType = (RestrictionType)tmpArray[0];
            switch (restrictionType)
            {
                case RestrictionType.AndRestriction:
                    this.Restriction = new AndRestriction(this.CountType);
                    break;
                case RestrictionType.BitMaskRestriction:
                    this.Restriction = new BitMaskRestriction();
                    break;
                case RestrictionType.CommentRestriction:
                    this.Restriction = new CommentRestriction(this.CountType);
                    break;
                case RestrictionType.ComparePropertiesRestriction:
                    this.Restriction = new ComparePropertiesRestriction();
                    break;
                case RestrictionType.ContentRestriction:
                    this.Restriction = new ContentRestriction();
                    break;
                case RestrictionType.CountRestriction:
                    this.Restriction = new CountRestriction(this.CountType);
                    break;
                case RestrictionType.ExistRestriction:
                    this.Restriction = new ExistRestriction();
                    break;
                case RestrictionType.NotRestriction:
                    this.Restriction = new NotRestriction(this.CountType);
                    break;
                case RestrictionType.OrRestriction:
                    this.Restriction = new OrRestriction(this.CountType);
                    break;
                case RestrictionType.PropertyRestriction:
                    this.Restriction = new PropertyRestriction();
                    break;
                case RestrictionType.SizeRestriction:
                    this.Restriction = new SizeRestriction();
                    break;
                case RestrictionType.SubObjectRestriction:
                    this.Restriction = new SubObjectRestriction(this.CountType);
                    break;
            }

            size += this.Restriction.Deserialize(tmpArray);
            return size;
        }
    }
}
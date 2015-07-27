//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXORULE
{
    #region Enums
    /// <summary>
    /// Enum of RestrictType used to create different Restriction
    /// [MS-OXCDATA] 2.12
    /// </summary>
    public enum RestrictionType : byte
    {
        /// <summary>
        /// Logical AND operation applied to a list of subrestrictions.
        /// </summary>
        AndRestriction = 0x00,

        /// <summary>
        /// Logical OR operation applied to a list of subrestrictions.
        /// </summary>
        OrRestriction = 0x01,

        /// <summary>
        /// Logical NOT applied to a subrestriction.
        /// </summary>
        NotRestriction = 0x02,

        /// <summary>
        /// Search a property value for specific content.
        /// </summary>
        ContentRestriction = 0x03,

        /// <summary>
        /// Compare a property value to a particular value.
        /// </summary>
        PropertyRestriction = 0X04,

        /// <summary>
        /// Logical NOT applied to a subrestriction.
        /// </summary>
        ComparePropertiesRestriction = 0x05,

        /// <summary>
        /// Perform bitwise AND of a property value with a mask and compare to zero.
        /// </summary>
        BitMaskRestriction = 0x06,

        /// <summary>
        /// Compare the Size of a property value to a particular figure.
        /// </summary>
        SizeRestriction = 0X07,

        /// <summary>
        /// Test whether a property has a value.
        /// </summary>
        ExistRestriction = 0x08,

        /// <summary>
        /// Test whether any row of a message's attachment or recipient table satisfies a subrestriction.
        /// </summary>
        SubObjectRestriction = 0x09,

        /// <summary>
        /// Associates a comment with a subrestriction.
        /// </summary>
        CommentRestriction = 0x0A,

        /// <summary>
        /// Limits the number of matches returned from a subrestriction.
        /// </summary>
        CountRestriction = 0x0B
    }

    /// <summary>
    /// Relational operator that is used to compare value.
    /// </summary>
    public enum RelationalOperator : byte
    {
        /// <summary>
        /// TRUE if the value of the object's property is less than the given value.
        /// </summary>
        RELOP_LT,

        /// <summary>
        /// TRUE if the value of the object's property is less than or equal to the given value.
        /// </summary>
        RELOP_LE,

        /// <summary>
        /// TRUE if the value of the object's property value is greater than the given value.
        /// </summary>
        RELOP_GT,

        /// <summary>
        /// TRUE if the value of the object's property value is greater than or equal to the given value.
        /// </summary>
        RELOP_GE,

        /// <summary>
        /// TRUE if the object's property value equals the given value.
        /// </summary>
        RELOP_EQ,

        /// <summary>
        /// TRUE if the object's property value does not equal the given value.
        /// </summary>
        RELOP_NE,

        /// <summary>
        /// TRUE if the value of the object's property is in the DL membership of the specified property value. The value of the object's property MUST be an EntryID of a mail-enabled object in the address book. The specified property value MUST be an EntryID of a distribution list object in the address book.
        /// </summary>
        RELOP_MEMBER_OF_DL
    }
    #endregion

    #region Restrictions
    /// <summary>
    /// Base class of Restrictions
    /// </summary>
    public abstract class Restrictions : IRestriction
    {
        /// <summary>
        /// Unsigned 8-bit integer. This value indicates the Type of restriction.
        /// </summary>
        private RestrictionType restrictType;

        /// <summary>
        /// Type of COUNT
        /// </summary>
        private CountByte countType;

        /// <summary>
        /// Initializes a new instance of the Restrictions class.
        /// </summary>
        public Restrictions()
        {
            this.countType = CountByte.TwoBytesCount;
        }

        /// <summary>
        /// Initializes a new instance of the Restrictions class.
        /// </summary>
        /// <param name="countType">The COUNT Type of this class.</param>
        public Restrictions(CountByte countType)
        {
            this.countType = countType;
        }

        /// <summary>
        /// Gets or sets unsigned 8-bit integer. This value indicates the Type of restriction.
        /// </summary>
        public RestrictionType RestrictType
        {
            get { return this.restrictType; }
            protected set { this.restrictType = value; }
        }

        /// <summary>
        /// Gets or sets type of Count
        /// </summary>
        public CountByte CountType
        {
            get { return this.countType; }
            protected set { this.countType = value; } 
        }

        /// <summary>
        /// Get the total Size of Restriction
        /// </summary>
        /// <returns>The Size of Restriction buffer.</returns>
        public abstract int Size();

        /// <summary>
        /// Get serialized byte array for this struct
        /// </summary>
        /// <returns>Serialized byte array.</returns>
        public abstract byte[] Serialize();

        /// <summary>
        /// Deserialized byte array to a Restriction instance
        /// </summary>
        /// <param name="buffer">Byte array contain data of a Restriction instance.</param>
        /// <returns>Bytes count that deserialized in buffer.</returns>
        public abstract uint Deserialize(byte[] buffer);
    }
    #endregion
}
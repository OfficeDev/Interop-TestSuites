//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System.Collections.Generic;

    /// <summary>
    /// Contains a list of propValues.
    /// propList             = *PropValue
    /// </summary>
    public class PropList : SyntacticalBase
    {
        /// <summary>
        /// A list of PropValue objects.
        /// </summary>
        private List<PropValue> propValues;

        /// <summary>
        /// Initializes a new instance of the PropList class.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public PropList(FastTransferStream stream)
            : base(stream)
        {
        }

        /// <summary>
        /// Gets or sets the PropValue list.
        /// </summary>
        public List<PropValue> PropValues
        {
            get { return this.propValues; }
            set { this.propValues = value; }
        }

        /// <summary>
        /// Gets a value indicating whether the PropValue list contains meta-properties.
        /// </summary>
        public bool IsNoMetaPropertyContained
        {
            get
            {
                if (this.PropValues != null && this.PropValues.Count > 0)
                { 
                    List<uint> list = EnumHelper.GetEnumValues<uint>();
                    foreach (PropValue val in this.PropValues)
                    {
                        foreach (uint e in list)
                        {
                            if (val.PropType == (ushort)(e & 0xffff))
                            {
                                if (val.PropInfo.PropID == e >> 16)
                                {
                                    return false;
                                }
                            }
                        }
                    }
                }

                return true;
            }
        }

        /// <summary>
        /// Gets a value indicating whether contains PidTagIdsetDeleted.
        /// </summary>
        public bool HasPidTagIdsetDeleted
        {
            get 
            {
                return this.HasPropertyTag(0x67e5, 0x0102);
            }
        }

        /// <summary>
        /// Gets a value indicating whether contains PidTagIdsetExpired.
        /// </summary>
        public bool HasPidTagIdsetExpired
        {
            get 
            {
                return this.HasPropertyTag(0x6793, 0x0102);
            }
        }

        /// <summary>
        /// Gets a value indicating whether contains PidTagIdsetNoLongerInScope.
        /// </summary>
        public bool HasPidTagIdsetNoLongerInScope
        {
            get 
            {
                return this.HasPropertyTag(0x4021, 0x0102);
            }
        }

        /// <summary>
        /// Verify that a stream's current position contains a serialized propList.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        /// <returns>If the stream's current position contains 
        /// a serialized propList, return true, else false.</returns>
        public static bool Verify(FastTransferStream stream)
        {
            return PropValue.Verify(stream);
        }

        /// <summary>
        /// Indicates whether has specified property tag.
        /// </summary>
        /// <param name="id">A ushort value, the id of the property tag.</param>
        /// <param name="type">A ushort value, the type of the property tag.</param>
        /// <returns>If contains the property tag, returns true, otherwise returns false.</returns>
        public bool HasPropertyTag(ushort id, ushort type)
        {
            if (this.propValues != null && this.propValues.Count > 0)
            {
                foreach (PropValue p in this.propValues)
                {
                    if (p.PropType == type && p.PropInfo.PropID == id)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Indicates whether contains a type.
        /// </summary>
        /// <param name="type">A ushort value.</param>
        /// <returns>True if contains the specific type, false otherwise.</returns>
        public bool HasPropertyType(ushort type)
        { 
            if (this.PropValues != null && this.PropValues.Count > 0)
            {
                foreach (PropValue v in this.PropValues)
                {
                    if (v.PropType == type)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Indicates whether contains an id.
        /// </summary>
        /// <param name="id">A ushort value.</param>
        /// <returns>True if contains the specific type, false otherwise.</returns>
        public bool HasPropertyID(ushort id)
        {
            if (this.PropValues != null && this.PropValues.Count > 0)
            {
                foreach (PropValue v in this.PropValues)
                {
                    if (v.PropInfo.PropID == id)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Gets the value specified by a property tag.
        /// </summary>
        /// <param name="id">A ushort value, the id of the property tag.</param>
        /// <param name="type">A ushort value, the type of the property tag.</param>
        /// <returns>If contains the tagedProperty value return the value, otherwise return null.</returns>
        public object GetPropValue(ushort id, ushort type)
        {
            if (this.propValues != null && this.propValues.Count > 0)
            {
                foreach (PropValue p in this.propValues)
                {
                    if (p.PropType == type && p.PropInfo.PropID == id)
                    {
                        if (p is VarPropTypePropValue)
                        {
                            return (p as VarPropTypePropValue).ValueArray;
                        }
                        else if (p is FixedPropTypePropValue)
                        {
                            return (p as FixedPropTypePropValue).FixedValue;
                        }
                        else if (p is MvPropTypePropValue)
                        {
                            MvPropTypePropValue mvp = p as MvPropTypePropValue;
                            if (mvp.VarSizeValueList != null)
                            {
                                return mvp.VarSizeValueList;
                            }
                            else if (mvp.FixedSizeValue != null)
                            {
                                return mvp.FixedSizeValue;
                            }
                        }
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Deserialize fields from a FastTransferStream.
        /// </summary>
        /// <param name="stream">A FastTransferStream.</param>
        public override void Deserialize(FastTransferStream stream)
        {
            this.propValues = new List<PropValue>();
            while (PropValue.Verify(stream)
                && !MarkersHelper.IsEndMarkerExceptEcWarning(stream.VerifyUInt32()))
            {
                this.propValues.Add(PropValue.DeserializeFrom(stream) as PropValue);
            }

            if (SyntacticalBase.AllPropList != null)
            {
                SyntacticalBase.AllPropList.Add(this);
            }
        }
    }
}
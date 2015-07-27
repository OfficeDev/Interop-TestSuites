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
    /// Supply help functions for lexical enumerations.
    /// </summary>
    public class LexicalTypeHelper : EnumHelper
    {
        /// <summary>
        /// Contains fixedPropTypes.
        /// </summary>
        private static List<PropertyDataType> fixedTypes;

        /// <summary>
        /// Contains varPropTypes.
        /// </summary>
        private static List<PropertyDataType> varTypes;

        /// <summary>
        /// Contains mvPropTypes.
        /// </summary>
        private static List<PropertyDataType> mvtypes;

        /// <summary>
        /// Initializes static members of the LexicalTypeHelper class.
        /// </summary>
        static LexicalTypeHelper()
        {
            fixedTypes = new List<PropertyDataType>
            {
                PropertyDataType.PtypInteger16,
                PropertyDataType.PtypInteger32,
                PropertyDataType.PtypFloating32,
                PropertyDataType.PtypFloating64,
                PropertyDataType.PtypCurrency,
                PropertyDataType.PtypFloatingTime,
                PropertyDataType.PtypErrorCode,
                PropertyDataType.PtypBoolean,
                PropertyDataType.PtypInteger64,
                PropertyDataType.PtypTime,
                PropertyDataType.PtypGuid
            };

            varTypes = new List<PropertyDataType>
            {
                PropertyDataType.PtypString,
                PropertyDataType.PtypString8,
                PropertyDataType.PtypBinary,
                PropertyDataType.PtypServerId,
                PropertyDataType.PtypObject
            };

            mvtypes = new List<PropertyDataType>
            {
                PropertyDataType.PtypMultipleInteger16,
                PropertyDataType.PtypMultipleInteger32,
                PropertyDataType.PtypMultipleFloating32,
                PropertyDataType.PtypMultipleFloating64,
                PropertyDataType.PtypMultipleCurrency,
                PropertyDataType.PtypMultipleFloatingTime,
                PropertyDataType.PtypMultipleInteger64,
                PropertyDataType.PtypMultipleString,
                PropertyDataType.PtypMultipleString8,
                PropertyDataType.PtypMultipleTime,
                PropertyDataType.PtypMultipleGuid,
                PropertyDataType.PtypMultipleBinary
            };
        }

        /// <summary>
        /// Gets any type that has a fixed length.
        /// </summary>
        public static List<PropertyDataType> FixedTypes
        {
            get
            {
                return fixedTypes;
            }
        }

        /// <summary>
        /// Gets PtypString, PtypString8 and PtypBinary, PtypServerId, and PtypObject. 
        /// </summary>
        public static List<PropertyDataType> VarTypes
        {
            get
            {
                return varTypes;
            }
        }

        /// <summary>
        /// Gets multi-valued property types.
        /// </summary>
        public static List<PropertyDataType> MVTypes
        {
            get
            {
                return mvtypes;
            }
        }

        /// <summary>
        /// Indicate whether a PropertyDataType is a multi-valued property type.
        /// </summary>
        /// <param name="type">A PropertyDataType.</param>
        /// <returns>If the PropertyDataType is a multi-value type 
        /// return true, else false.</returns>
        public static bool IsMVType(PropertyDataType type)
        {
            return mvtypes.Contains(type);
        }

        /// <summary>
        /// Indicate whether a PropertyDataType is either 
        /// PtypString, PtypString8 or PtypBinary, PtypServerId, or PtypObject. 
        /// </summary>
        /// <param name="type">A PropertyDataType.</param>
        /// <returns>If the PropertyDataType is a either 
        /// PtypString, PtypString8 or PtypBinary, PtypServerId, or PtypObject 
        /// return true, else false.
        /// </returns>
        public static bool IsVarType(PropertyDataType type)
        {
            return varTypes.Contains(type);
        }

        /// <summary>
        /// Indicate whether a property type value of any type that has a fixed length.
        /// </summary>
        /// <param name="type">A property type.</param>
        /// <returns>If a property type value of any type that has a fixed length,
        /// return true , else return false.
        /// </returns>
        public static bool IsFixedType(PropertyDataType type)
        {
            return fixedTypes.Contains(type);
        }

        /// <summary>
        /// Indicate whether a PropertyID is a named property ID.
        /// </summary>
        /// <param name="id">A ushort value.</param>
        /// <returns>If a PropertyID is a named property ID,
        /// return true, else return false.
        /// </returns>
        public static bool IsNamedPropertyID(ushort id)
        {
            return id >= 0x8000;
        }

        /// <summary>
        /// Indicate whether a PropertyID is a tagged property ID.
        /// </summary>
        /// <param name="id">A ushort value.</param>
        /// <returns>If a PropertyID is a tagged property ID,
        /// return true, else return false.
        /// </returns>
        public static bool IsTaggedPropertyID(ushort id)
        {
            return id < 0x8000;
        }
    }
}
//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXCFOLD
{
    using System;

    /// <summary>
    /// Deserialize the common restricts type.
    /// </summary>
    public class RestrictsFactory
    {
        /// <summary>
        /// Deserialize the common field of restriction.
        /// </summary>
        /// <param name="restrictionData">The restriction data.</param>
        /// <returns>Return a new instance of a class extended from Restricts.</returns>
        public static Restriction Deserialize(byte[] restrictionData)
        {
            Restriction restricts = null;
            int index = 0;
            byte restrictType = restrictionData[index];
            switch (restrictType)
            {
                case (byte)RestrictType.AndRestriction:
                    AndRestriction andRestriction = new AndRestriction();
                    andRestriction.Deserialize(restrictionData);
                    restricts = andRestriction;
                    break;
                case (byte)RestrictType.NotRestriction:
                    NotRestriction notRestriction = new NotRestriction();
                    notRestriction.Deserialize(restrictionData);
                    restricts = notRestriction;
                    break;
                case (byte)RestrictType.ContentRestriction:
                    ContentRestriction contentRestriction = new ContentRestriction();
                    contentRestriction.Deserialize(restrictionData);
                    restricts = contentRestriction;
                    break;
                case (byte)RestrictType.PropertyRestriction:
                    PropertyRestriction propertyRestriction = new PropertyRestriction();
                    propertyRestriction.Deserialize(restrictionData);
                    restricts = propertyRestriction;
                    break;
                case (byte)RestrictType.ExistRestriction:
                    ExistRestriction existRestriction = new ExistRestriction();
                    existRestriction.Deserialize(restrictionData);
                    restricts = existRestriction;
                    break;
                default:
                    throw new NotSupportedException("This restrict type '" + restrictType.ToString() + "'is not supported in this test suite.");
            }

            return restricts;
        }
    }
}
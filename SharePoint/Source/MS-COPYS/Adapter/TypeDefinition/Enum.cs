//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_COPYS
{
    /// <summary>
    /// The enum represents the destination type.
    /// </summary>
    public enum DestinationFileUrlType
    {  
       /// <summary>
       /// Represents the destination is a document library under normal site on destination SUT.(The destination document library and source document library are in different SUT)
       /// </summary>
       NormalDesLibraryOnDesSUT = 0,

       /// <summary>
       /// Represents the destination is a document library under Meeting Work Space on destination SUT.(The destination document library and source document library are in different SUT)
       /// </summary>
       MWSLibraryOnDestinationSUT = 1,
    }

    /// <summary>
    /// The enum specifies the source file URL type.
    /// </summary>
    public enum SourceFileUrlType
    {  
       /// <summary>
       /// Represents the source file URL stored on source SUT. 
       /// </summary>
       SourceFileOnSourceSUT = 0,

       /// <summary>
       ///  Represents the source file URL stored on destination SUT. 
       /// </summary>
       SourceFileOnDesSUT = 1
    }

    /// <summary>
    /// The enum specifies the protocol SUT to connect. 
    /// </summary>
    public enum ServiceLocation
    {
        /// <summary>
        /// Represents the Source SUT.
        /// </summary>
        SourceSUT = 0,

        /// <summary>
        /// Represents the destination SUT.
        /// </summary>
        DestinationSUT = 1,
    }

    /// <summary>
    /// The enum specifies the field attribute types. It indicates which attribute will be used.
    /// </summary>
    public enum FieldAttributeType
    {   
        /// <summary>
        /// Represents the DisplayName attribute of field.
        /// </summary>
        DisplayName = 0,

        /// <summary>
        /// Represents the InternalName attribute of field.
        /// </summary>
        InternalName = 1,

        /// <summary>
        /// Represents the Id attribute of field.
        /// </summary>
        Id = 2,
    }
}
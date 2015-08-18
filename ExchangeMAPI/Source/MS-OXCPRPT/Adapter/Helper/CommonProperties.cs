namespace Microsoft.Protocols.TestSuites.MS_OXCPRPT
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Common properties class.
    /// </summary>
    public partial class MS_OXCPRPTAdapter : IMS_OXCPRPTAdapter
    {
        /// <summary>
        /// Common properties.
        /// </summary>
        private Dictionary<CommonObjectProperty, PropertyTag> commonProperties;

        /// <summary>
        /// Initialize dictionary of common properties.
        /// </summary>
        private void InitializeCommonProperties()
        { 
            PropertyTag pidTagAccess;
            pidTagAccess.PropertyId = (ushort)TaggedPropertyName.PidTagAccess;
            pidTagAccess.PropertyType = (ushort)PropertyTypeName.PtypInteger32;

            PropertyTag pidTagAccessLevel;
            pidTagAccessLevel.PropertyId = (ushort)TaggedPropertyName.PidTagAccessLevel;
            pidTagAccessLevel.PropertyType = (ushort)PropertyTypeName.PtypInteger32;

            PropertyTag pidTagChangeKey;
            pidTagChangeKey.PropertyId = (ushort)TaggedPropertyName.PidTagChangeKey;
            pidTagChangeKey.PropertyType = (ushort)PropertyTypeName.PtypBinary;

            PropertyTag pidTagCreationTime;
            pidTagCreationTime.PropertyId = (ushort)TaggedPropertyName.PidTagCreationTime;
            pidTagCreationTime.PropertyType = (ushort)PropertyTypeName.PtypTime;

            PropertyTag pidTagLastModifierName;
            pidTagLastModifierName.PropertyId = (ushort)TaggedPropertyName.PidTagLastModifierName;
            pidTagLastModifierName.PropertyType = (ushort)PropertyTypeName.PtypString;

            PropertyTag pidTagLastModificationTime;
            pidTagLastModificationTime.PropertyId = (ushort)TaggedPropertyName.PidTagLastModificationTime;
            pidTagLastModificationTime.PropertyType = (ushort)PropertyTypeName.PtypTime;

            PropertyTag pidTagObjectType;
            pidTagObjectType.PropertyId = (ushort)TaggedPropertyName.PidTagObjectType;
            pidTagObjectType.PropertyType = (ushort)PropertyTypeName.PtypInteger32;

            PropertyTag pidTagRecordKey;
            pidTagRecordKey.PropertyId = (ushort)TaggedPropertyName.PidTagRecordKey;
            pidTagRecordKey.PropertyType = (ushort)PropertyTypeName.PtypBinary;

            PropertyTag pidTagSearchKey;
            pidTagSearchKey.PropertyId = (ushort)TaggedPropertyName.PidTagSearchKey;
            pidTagSearchKey.PropertyType = (ushort)PropertyTypeName.PtypBinary;

            PropertyTag pidTagDisplayName;
            pidTagDisplayName.PropertyId = (ushort)TaggedPropertyName.PidTagDisplayName;
            pidTagDisplayName.PropertyType = (ushort)PropertyTypeName.PtypString;

            PropertyTag pidTagFolderId;
            pidTagFolderId.PropertyId = (ushort)TaggedPropertyName.PidTagFolderId;
            pidTagFolderId.PropertyType = (ushort)PropertyTypeName.PtypInteger64;

            this.commonProperties = new Dictionary<CommonObjectProperty, PropertyTag>
            {
                {
                    CommonObjectProperty.PidTagDisplayName, pidTagDisplayName
                },
                {
                    CommonObjectProperty.PidTagFolderId, pidTagFolderId
                },
                {
                    CommonObjectProperty.PidTagAccess, pidTagAccess
                },
                {
                    CommonObjectProperty.PidTagAccessLevel, pidTagAccessLevel
                },
                {
                    CommonObjectProperty.PidTagChangeKey, pidTagChangeKey
                },
                {
                    CommonObjectProperty.PidTagCreationTime, pidTagCreationTime
                },
                {
                    CommonObjectProperty.PidTagLastModificationTime, pidTagLastModificationTime
                },
                {
                    CommonObjectProperty.PidTagLastModifierName, pidTagLastModifierName
                },
                {
                    CommonObjectProperty.PidTagObjectType, pidTagObjectType
                },
                {
                    CommonObjectProperty.PidTagRecordKey, pidTagRecordKey
                },
                {
                    CommonObjectProperty.PidTagSearchKey, pidTagSearchKey
                }
            };
        }
    }
}
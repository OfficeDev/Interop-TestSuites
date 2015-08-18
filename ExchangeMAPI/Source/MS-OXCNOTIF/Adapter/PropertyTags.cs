namespace Microsoft.Protocols.TestSuites.MS_OXCNOTIF
{
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;

    /// <summary>
    /// Property Tags
    /// </summary>
    public class PropertyTags : Dictionary<PropertyNames, PropertyTag>
    {
        /// <summary>
        /// PropertyTags All.
        /// </summary>        
        private static PropertyTags all;

        /// <summary>
        /// Gets PropertyTags all.
        /// </summary>
        public static PropertyTags All
        {
            get
            {
                if (all == null)
                {
                    all = new PropertyTags
                    {
                        {
                            PropertyNames.PidTagHasAttachments, 0x0E1B, PropertyType.PtypBoolean
                        },
                        {
                            PropertyNames.PidTagMessageClass, 0x001A, PropertyType.PtypString8
                        },
                        {
                            PropertyNames.PidTagMessageCodepage, 0x3FFD, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagMessageLocaleId, 0x3FF1, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagMessageSize, 0x0E08, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagMessageStatus, 0x0E17, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagSubjectPrefix, 0x003D, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagNormalizedSubject, 0x0E1D, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagImportance, 0x0017, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagPriority, 0x0026, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagSensitivity, 0x0036, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidLidSmartNoAttach, 0x00008514, PropertyType.PtypBoolean
                        },
                        {
                            PropertyNames.PidLidPrivate, 0x00008506, PropertyType.PtypBoolean
                        },
                        {
                            PropertyNames.PidLidSideEffects, 0x00008510, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidNameKeywords, 0x0000, PropertyType.PtypMultipleString
                        },
                        {
                            PropertyNames.PidLidCommonStart, 0x00008516, PropertyType.PtypTime
                        },
                        {
                            PropertyNames.PidLidCommonEnd, 0x00008517, PropertyType.PtypTime
                        },
                        {
                            PropertyNames.PidTagAutoForwardComment, 0x0004, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagBody, 0x1000, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagNativeBody, 0x1016, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagBodyHtml, 0x1013, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagRtfCompressed, 0x1009, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagRtfInSync, 0x0E1F, PropertyType.PtypBoolean
                        },
                        {
                            PropertyNames.PidTagInternetCodepage, 0x3FDE, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidLidContactLinkEntry, 0x00008585, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidLidContacts, 0x0000853A, PropertyType.PtypMultipleString
                        },
                        {
                            PropertyNames.PidLidContactLinkName, 0x00008586, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidLidContactLinkSearchKey, 0x00008584, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagArchiveTag, 0x3018, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagPolicyTag, 0x3019, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagRetentionPeriod, 0x301A, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagStartDateEtc, 0x301B, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagRetentionDate, 0x301C, PropertyType.PtypTime
                        },
                        {
                            PropertyNames.PidTagRetentionFlags, 0x301D, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagArchivePeriod, 0x301E, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagArchiveDate, 0x301F, PropertyType.PtypTime
                        },
                        {
                            PropertyNames.PidTagLastModificationTime, 0x3008, PropertyType.PtypTime
                        },
                        {
                            PropertyNames.PidTagCreationTime, 0x3007, PropertyType.PtypTime
                        },
                        {
                            PropertyNames.PidTagDisplayName, 0x3001, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagAttachSize, 0x0E20, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagAttachNumber, 0x0E21, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagAttachPathname, 0x3708, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagAttachDataBinary, 0x3701, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagAttachMethod, 0x3705, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagAttachFilename, 0x3704, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagAttachLongFilename, 0x3707, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagAttachExtension, 0x3703, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagAttachLongPathname, 0x370D, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagAttachTag, 0x370A, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagRenderingPosition, 0X370B, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagAttachRendering, 0x3709, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagAttachFlags, 0x3714, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagAttachTransportName, 0x370C, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagAttachEncoding, 0x3702, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagAttachAdditionalInformation, 0x370F, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagAttachmentLinkId, 0x7FFA, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagAttachmentFlags, 0x7FFD, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagAttachmentHidden, 0x7FFE, PropertyType.PtypBoolean
                        },
                        {
                            PropertyNames.PidTagAttachMimeTag, 0x370E, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagAttachContentId, 0x3712, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagAttachContentLocation, 0x3713, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagAttachContentBase, 0x3711, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagAttachDataObject, 0x3701, PropertyType.PtypComObject
                        },
                        {
                            PropertyNames.PidTagMessageFlags, 0x0E07, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagAccessLevel, 0x0FF7, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagDisplayBcc, 0x0E02, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagDisplayCc, 0x0E03, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagDisplayTo, 0x0E04, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagSecurityDescriptor, 0x0E27, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagUrlCompNameSet, 0x0E62, PropertyType.PtypBoolean
                        },
                        {
                            PropertyNames.PidTagTrustSender, 0x0E79, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagUrlCompName, 0x10F3, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagSearchKey, 0x300B, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagAccess, 0x0FF4, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagCreatorName, 0x3FF8, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagCreatorEntryId, 0x3FF9, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagLastModifierName, 0x3FFA, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagLastModifierEntryId, 0x3FFB, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagHasNamedProperties, 0x664A, PropertyType.PtypBoolean
                        },
                        {
                            PropertyNames.PidTagLocaleId, 0x66A1, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagChangeKey, 0x65E2, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagObjectType, 0x0FFE, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagRecordKey, 0x0FF9, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagLocalCommitTime, 0x6709, PropertyType.PtypTime
                        },
                        {
                            PropertyNames.PidTagAutoForwarded, 0x0005, PropertyType.PtypBoolean
                        },
                        {
                            PropertyNames.PidTagDisplayType, 0x3900, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagAddressBookDisplayNamePrintable, 0x39ff, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagSmtpAddress, 0x39fe, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagSendInternetEncoding, 0x3a71, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagDisplayTypeEx, 0x3905, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagRecipientDisplayName, 0x5ff6, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagRecipientFlags, 0x5ffd, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagRecipientTrackStatus, 0x5fff, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagRecipientResourceState, 0x5fde, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagRecipientOrder, 0x5fdf, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagRecipientEntryId, 0x5ff7, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagFolderId, 0x6748, PropertyType.PtypInteger64
                        },
                        {
                            PropertyNames.PidTagMid, 0x674A, PropertyType.PtypInteger64
                        },
                        {
                            PropertyNames.PidTagInstID, 0x674D, PropertyType.PtypInteger64
                        },
                        {
                            PropertyNames.PidTagInstanceNum, 0x674E, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagSubject, 0x0037, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagMessageDeliveryTime, 0x0E06, PropertyType.PtypTime
                        },
                        {
                            PropertyNames.PidTagRowType, 0x0FF5, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagContentCount, 0x3602, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagOfflineAddressBookName, 0x6800, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagOfflineAddressBookSequence, 0x6801, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagOfflineAddressBookContainerGuid, 0x6802, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagOfflineAddressBookMessageClass, 0x6803, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagOfflineAddressBookDistinguishedName, 0x6804, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagSortLocaleId, 0x6705, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagEntryId, 0x0FFF, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagMemberId, 0x6671, PropertyType.PtypInteger64
                        },
                        {
                            PropertyNames.PidTagMemberName, 0x6672, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagMemberRights, 0x6673, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagRuleSequence, 0x6676, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagRuleCondition, 0x6679, PropertyType.PtypRestriction
                        },
                        {
                            PropertyNames.PidTagRuleActions, 0x6680, PropertyType.PtypRuleAction
                        },
                        {
                            PropertyNames.PidTagRuleProvider, 0x6681, PropertyType.PtypString
                        },
                        {
                            PropertyNames.PidTagCnsetSeen, 0x6796, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagSourceKey, 0x65E0, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagPredecessorChangeList, 0x65E3, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagParentSourceKey, 0x65E1, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.PidTagFolderType, 0x3601, PropertyType.PtypInteger32
                        },
                        {
                            PropertyNames.PidTagTemplateData, 0x0001, PropertyType.PtypBinary
                        },
                        {
                            PropertyNames.UserSpecified, 0x1234, PropertyType.PtypString
                        }
                    };
                }

                return PropertyTags.all;
            }
        }

        /// <summary>
        /// Initialize the properties into the dictionary. 
        /// </summary>
        /// <param name="propertyName">The name of property</param>
        /// <param name="propertyId">Property ID</param>
        /// <param name="propertyType">Property type ID</param>
        private void Add(PropertyNames propertyName, ushort propertyId, PropertyType propertyType)
        {
            PropertyTag propertyTag = new PropertyTag
            {
                PropertyId = propertyId, PropertyType = (ushort)propertyType
            };
            this.Add(propertyName, propertyTag);
        }
    }
}
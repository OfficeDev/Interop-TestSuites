//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OUTSPS
{
    using System;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This partial class is the container for the adapter capture code.
    /// </summary>
    public partial class MS_OUTSPSAdapter : ManagedAdapterBase, IMS_OUTSPSAdapter
    {
        /// <summary>
        /// A method used to verify the requirements for common schema definition of list.
        /// </summary>
        /// <param name="responseOfGetList">A parameter represents the response of GetList operation which contain common fields' definitions.</param>
        private void VerifyCommonSchemaOfListDefinition(GetListResponseGetListResult responseOfGetList)
        {
           if (null == responseOfGetList)
           {
               throw new ArgumentNullException("responseOfGetList");
           }

           // Verify Attachments field's id and type.
           string actuallistTemplateValue = responseOfGetList.List.ServerTemplate;
           string docLibraryListTemplateStringValue = ((int)TemplateType.Document_Library).ToString();
           if (!docLibraryListTemplateStringValue.Equals(actuallistTemplateValue, StringComparison.OrdinalIgnoreCase))
           {
               bool isVerifyR584 = this.VerifyTypeAndIdForFieldDefinition(
                                     responseOfGetList,
                                     "Attachments",
                                     "{67df98f4-9dec-48ff-a553-29bece9c5bf4}",
                                     "Attachments");

               // Verify MS-OUTSPS requirement: MS-OUTSPS_R584
               this.Site.CaptureRequirementIfIsTrue(
                                               isVerifyR584,
                                               584,
                                               @"[In Common Schema][One of the common properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Attachments[ Field.ID:]{67df98f4-9dec-48ff-a553-29bece9c5bf4}[Field.Type:]Attachments.");
           }
     
           // Verify ContentTypeId field's id and type.
           bool isVerifyR586 = this.VerifyTypeAndIdForFieldDefinition(
                                           responseOfGetList,
                                           "ContentTypeId",
                                           "{03e45e84-1992-4d42-9116-26f756012634}",
                                           "ContentTypeId");

           // Verify MS-OUTSPS requirement: MS-OUTSPS_R586
           this.Site.CaptureRequirementIfIsTrue(
                                           isVerifyR586,
                                           586,
                                           @"[In Common Schema][One of the common properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]ContentTypeId[ Field.ID:]{03e45e84-1992-4d42-9116-26f756012634}[Field.Type:]ContentTypeId.");

           // Verify Created field's id and type.
           bool isVerifyR587 = this.VerifyTypeAndIdForFieldDefinition(
                                           responseOfGetList,
                                           "Created",
                                           "{8c06beca-0777-48f7-91c7-6da68bc07b69}",
                                           "DateTime");

           // Verify MS-OUTSPS requirement: MS-OUTSPS_R587
           this.Site.CaptureRequirementIfIsTrue(
                                           isVerifyR587,
                                           587,
                                           @"[In Common Schema][One of the common properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Created[ Field.ID:]{8c06beca-0777-48f7-91c7-6da68bc07b69}[Field.Type:]DateTime.");

           // Verify ID field's id and type.
           bool isVerifyR588 = this.VerifyTypeAndIdForFieldDefinition(
                                           responseOfGetList,
                                           "ID",
                                           "{1d22ea11-1e32-424e-89ab-9fedbadb6ce1}",
                                           "Counter");

           // Verify MS-OUTSPS requirement: MS-OUTSPS_R588
           this.Site.CaptureRequirementIfIsTrue(
                                           isVerifyR588,
                                           588,
                                           @"[In Common Schema][One of the common properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]ID[ Field.ID:]{1d22ea11-1e32-424e-89ab-9fedbadb6ce1}[Field.Type:]Counter.");

           // Verify Modified field's id and type.
           bool isVerifyR589 = this.VerifyTypeAndIdForFieldDefinition(
                                           responseOfGetList,
                                           "Modified",
                                           "{28cf69c5-fa48-462a-b5cd-27b6f9d2bd5f}",
                                           "DateTime");

           // Verify MS-OUTSPS requirement: MS-OUTSPS_R589
           this.Site.CaptureRequirementIfIsTrue(
                                           isVerifyR589,
                                           589,
                                           @"[In Common Schema][One of the common properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Modified[ Field.ID:]{28cf69c5-fa48-462a-b5cd-27b6f9d2bd5f}[Field.Type:]DateTime.");

           // Verify owshiddenversion field's id and type.
           bool isVerifyR590 = this.VerifyTypeAndIdForFieldDefinition(
                                           responseOfGetList,
                                           "owshiddenversion",
                                           "{d4e44a66-ee3a-4d02-88c9-4ec5ff3f4cd5}",
                                           "Integer");

           // Verify MS-OUTSPS requirement: MS-OUTSPS_R590
           this.Site.CaptureRequirementIfIsTrue(
                                           isVerifyR590,
                                           590,
                                           @"[In Common Schema][One of the common properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]owshiddenversion[ Field.ID:]{d4e44a66-ee3a-4d02-88c9-4ec5ff3f4cd5}[Field.Type:]Integer.");
        }

        /// <summary>
        /// A method used to verify response of GetGetListItemChangesSinceToken operation and capture related requirements.
        /// </summary>
        /// <param name="responseOfGetGetListItemChangesSinceToken">A parameter represents the response of GetGetListItemChangesSinceToken operation where this method perform the verification.</param>
        private void VerifyGetListItemChangesSinceTokenResponse(GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResult responseOfGetGetListItemChangesSinceToken)
        {
            if (null == responseOfGetGetListItemChangesSinceToken)
            {
                throw new ArgumentNullException("responseOfGetGetListItemChangesSinceToken");
            }

            string formatedErrorMessage = @"The [{0}] should present in response of GetListItemChangesSinceToken operation.";
            this.Site.Assert.IsNotNull(
                                responseOfGetGetListItemChangesSinceToken.listitems,
                                formatedErrorMessage,
                                "listitems");
            GetListItemChangesSinceTokenResponseGetListItemChangesSinceTokenResultListitems listitemsIntance = responseOfGetGetListItemChangesSinceToken.listitems;

            // If the Changes.Id element is present in the response, it means there are special changes for stored data of list items, so that the protocol SUT does not query the information for list items, such as the list items are in Restore, Move, InvalidToken, Delete. The code return now since the expected data needed below will not be present in response.
            if (listitemsIntance.Changes.Id != null && listitemsIntance.Changes.Id.ChangeTypeSpecified)
            {
                return;
            }

            // If the protocol SUT query the valid changes and put them in the response and the AlternateUrls is not null or empty, then capture R1222
            this.Site.Assert.IsFalse(
                                    string.IsNullOrEmpty(listitemsIntance.AlternateUrls),
                                    formatedErrorMessage,
                                    "AlternateUrls");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1222
            this.Site.CaptureRequirement(
                                    1222,
                                    @"[In GetListItemChangesSinceTokenResponse][The element]AlternateUrl is contained within GetListItemChangesSinceTokenResponse.GetListItemChangesSinceTokenResult.listitems, as specified by [MS-LISTSWS].");

            // If there are any list items change data, then perform the verification. 
            if (!listitemsIntance.data.ItemCount.Equals("0"))
            {
                // If the EffectivePermMask is not null or empty, then capture R1227
                this.Site.Assert.IsFalse(
                                      string.IsNullOrEmpty(listitemsIntance.EffectivePermMask),
                                      formatedErrorMessage,
                                      "EffectivePermMask");

                // Verify MS-OUTSPS requirement: MS-OUTSPS_R1227
                this.Site.CaptureRequirement(
                                        1227,
                                        @"[In GetListItemChangesSinceTokenResponse][The element]EffectivePermMask is contained within GetListItemChangesSinceTokenResponse.GetListItemChangesSinceTokenResult.listitems, as specified by [MS-LISTSWS].");
            }
         
            // If the MaxBulkDocumentSyncSize is not null or empty, then capture R1228
            this.Site.Assert.IsTrue(
                                  listitemsIntance.MaxBulkDocumentSyncSizeSpecified,
                                  formatedErrorMessage,
                                  "MaxBulkDocumentSyncSize");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1228
            this.Site.CaptureRequirement(
                                    1228,
                                    @"[In GetListItemChangesSinceTokenResponse][The element]MaxBulkDocumentSyncSize is contained within GetListItemChangesSinceTokenResponse.GetListItemChangesSinceTokenResult.listitems, as specified by [MS-LISTSWS].");

            // If the MinTimeBetweenSyncs is not null or empty, then capture R1229
            this.Site.Assert.IsTrue(
                                  listitemsIntance.MinTimeBetweenSyncsSpecified,
                                  formatedErrorMessage,
                                  "MinTimeBetweenSyncs");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1229
            this.Site.CaptureRequirement(
                                   1229,
                                   @"[In GetListItemChangesSinceTokenResponse][The element]MinTimeBetweenSyncs is contained within GetListItemChangesSinceTokenResponse.GetListItemChangesSinceTokenResult.listitems, as specified by [MS-LISTSWS].");

            // If the RecommendedTimeBetweenSyncs is not null or empty, then capture R1230
            this.Site.Assert.IsTrue(
                                  listitemsIntance.RecommendedTimeBetweenSyncsSpecified,
                                  formatedErrorMessage,
                                  "RecommendedTimeBetweenSyncs");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1230
            this.Site.CaptureRequirement(
                                   1230,
                                   @"[In GetListItemChangesSinceTokenResponse][The element]RecommendedTimeBetweenSyncs is contained within GetListItemChangesSinceTokenResponse.GetListItemChangesSinceTokenResult.listitems, as specified by [MS-LISTSWS].");

            // Verify whether the changes element have instance.
            this.Site.Assert.IsNotNull(
                                   listitemsIntance.Changes,
                                    formatedErrorMessage,
                                    "Changes");

            // If the Changes.LastChangeToken is not null, then capture R1224.
            this.Site.Assert.IsFalse(
                                string.IsNullOrEmpty(listitemsIntance.Changes.LastChangeToken),
                                formatedErrorMessage,
                                "LastChangeToken");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1224
            this.Site.CaptureRequirement(
                                    1224,
                                    @"[In GetListItemChangesSinceTokenResponse][The attribute]Changes.LastChangeToken is contained within GetListItemChangesSinceTokenResponse.GetListItemChangesSinceTokenResult.listitems, as specified by [MS-LISTSWS].");

            // If upon verification pass, then capture R1221
            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1221
            this.Site.CaptureRequirement(
                                    1221,
                                    @"[In GetListItemChangesSinceTokenResponse]Each of these[AlternateUrls,Changes.Id.ChangeType,Changes.data.ListItemCollectionPositionNext,Changes.LastChangeToken,Changes.MoreChanges,EffectivePermMask,MaxBulkDocumentSyncSize,MinTimeBetweenSyncs,RecommendedTimeBetweenSyncs] is contained within GetListItemChangesSinceTokenResponse.GetListItemChangesSinceTokenResult.listitems, as specified by [MS-LISTSWS].");
        }

        /// <summary>
        /// A method used to capture transport related requirements according to the current transport protocol the test suite use.
        /// </summary>
        private void VerifyTransportRequirement()
        {
            TransportProtocol currentTransportProtocol = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            if (TransportProtocol.HTTP == currentTransportProtocol)
            {
                // Verify MS-OUTSPS requirement: MS-OUTSPS_R868
                this.Site.CaptureRequirement(
                    868,
                    @"[In Transport] Protocol servers MUST support SOAP over HTTP.");
            }
            else if (TransportProtocol.HTTPS == currentTransportProtocol)
            {  
                if (Common.IsRequirementEnabled(892, this.Site))
                {
                    // Verify MS-OUTSPS requirement: MS-OUTSPS_R892
                    this.Site.CaptureRequirement(
                        892,
                        @"[In Appendix B: Product Behavior]Implementation does additionally support SOAP over HTTPS for securing communication with clients. (Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
                }
            }

            // If the SOAP operations communicate with protocol SUT, then capture R870
            this.Site.CaptureRequirement(
                                870,
                                @"[In Transport] This protocol uses the same transport, security model, and SOAP versions as the Lists Web Service Protocol ([MS-LISTSWS]).");
        }

        /// <summary>
        /// A method used to verify field's id or field's type is equal to expected value.
        /// </summary>
        /// <param name="responseOfGetList">A parameter represents the response of GetList operation which contains the field definitions.</param>
        /// <param name="expectedFieldName">A parameter represents the name of a field definition which is used to get the definition from the  response of GetList operation.</param>
        /// <param name="expectedFieldId">A parameter represents the expected id of field definition.</param>
        /// <param name="expectedFieldType">A parameter represents the expected type of field definition.</param>
        /// <returns>Return true indicating the verification pass.</returns>
        private bool VerifyTypeAndIdForFieldDefinition(GetListResponseGetListResult responseOfGetList, string expectedFieldName, string expectedFieldId, string expectedFieldType)
        {
            if (string.IsNullOrEmpty(expectedFieldName))
            {
                throw new ArgumentNullException("expectedFieldName");
            }

            // Get the field definition from response.
            FieldDefinition fieldDefinition = Common.GetFieldItemByName(responseOfGetList, expectedFieldName, this.Site);
            
            // Verify field type.
            this.Site.Assert.IsTrue(
                                    Common.VerifyFieldType(fieldDefinition, expectedFieldType, this.Site),
                                    @"The field definition's type should match the specified value in protocol.");

            // Verify field id.
            this.Site.Assert.IsTrue(
                                    Common.VerifyFieldId(fieldDefinition, expectedFieldId, this.Site),
                                    @"The field definition's id should match the specified value in protocol.");

            // If upon verifications pass, return true.
            return true;
        }   
    }
}
//-----------------------------------------------------------------------
// Copyright (c) 2015 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_OXWSATT
{
    using System;
    using System.Collections.Generic;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// The base class of scenario class.
    /// </summary>
    public class TestSuiteBase : TestClassBase
    {
        #region Properties

        /// <summary>
        /// Gets the MS-OXWSATT protocol adapter interface.
        /// </summary>
        protected IMS_OXWSATTAdapter ATTAdapter { get; private set; }

        /// <summary>
        /// Gets the MS-OXWSCORE protocol adapter interface.
        /// </summary>
        protected IMS_OXWSCOREAdapter COREAdapter { get; private set; }

        /// <summary>
        /// Gets the id of the message which attachments will attach to.
        /// </summary>
        protected ItemIdType ItemId { get; private set; }

        #endregion

        #region Test case initialize and clean up
        /// <summary>
        /// Initialize the Test suite.
        /// </summary>
        protected override void TestInitialize()
        {
            base.TestInitialize();

            this.ATTAdapter = this.Site.GetAdapter<IMS_OXWSATTAdapter>();
            this.COREAdapter = this.Site.GetAdapter<IMS_OXWSCOREAdapter>();

            // Create an item.
            this.ItemId = this.CreateMessage();
        }

        /// <summary>
        /// Clean up the environment.
        /// </summary>
        protected override void TestCleanup()
        {
            this.DeleteMessage(this.ItemId.Id);

            // Reset parameters.
            this.ItemId = null;

            base.TestCleanup();
        }
        #endregion

        #region Test case base methods

        /// <summary>
        /// Creates an item or file attachment on an item.
        /// </summary>
        /// <param name="parentItemId">Identifies the parent item in the server store that contains the attachment.</param>
        /// <param name="attachmentsType">Attachment type.</param>
        /// <returns>A response message for "CreateAttachment" operation.</returns>
        protected CreateAttachmentResponseType CallCreateAttachmentOperation(string parentItemId, params AttachmentTypeValue[] attachmentsType)
        {
            // Configure attachments.
            int attachmentCount = attachmentsType.Length;
            AttachmentType[] attachments = new AttachmentType[attachmentCount];
            for (int attachmentIndex = 0; attachmentIndex < attachmentCount; attachmentIndex++)
            {
                AttachmentType attachment = null;

                if (attachmentsType[attachmentIndex] == AttachmentTypeValue.FileAttachment)
                {
                    FileAttachmentType fileAttachment = new FileAttachmentType()
                    {
                        ContentLocation = @"http://www.contoso.com/xyz.abc",
                        Name = "attach.jpg",

                        // Ensure content id is unique.
                        ContentId = Guid.NewGuid().ToString(),
                        ContentType = "image/jpeg",
                        Content = Convert.FromBase64String("/9j/4AAQSkZJRgABAQEAYABgAAD/7AARRHVja3kAAQAEAAAARgAA/9sAQwACAQECAQECAgICAgICAgMFAwMDAwMGBAQDBQcGBwcHBgcHCAkLCQgICggHBwoNCgoLDAwMDAcJDg8NDA4LDAwM/9sAQwECAgIDAwMGAwMGDAgHCAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwMDAwM/8AAEQgADAAUAwEiAAIRAQMRAf/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/aAAwDAQACEQMRAD8A5j9hnXPH/gy003RdI+KejWVjciO7muPF86y2GiWFu3nahceUJUyEtfOcKroplMZeSJAxH1j+2p8f9L8TfBTw1qvwo8YSSWGr6FB4p1bStWsLi01TVPDl4ZLe21CwkeOJ9yXf2cSQlCfKukdjDvhM34beLf2n/G3wz+B2g+ItA1qbS9Y0bxBm1uIlDFA9vPG6kNkFWQlSp4welZfxG/4K1fHj9ozQj4T8V+N73U9J1u7e7uxLmWU+YYWaGJ5Cxt7fzLeF/s8Hlxbox8mABXNQjVjQnQlVk+zurrRbWSXnbbuedTyvD0ISoQvZ333WnyPqLxp8PG8feIrjU5dG8O60ZWMYuDeW1tjyyUKeWzALtZWGFG3Oe+aK/OP4ja3c+CfiZ4m07TJTbWdvq90iIQJSAsrKMs+WJwo6miuFZRbT2svvX+QoZTQirXf4f5H/2Q==")
                    };

                    attachment = fileAttachment;
                }
                else
                {
                    ItemAttachmentType itemAttachment = new ItemAttachmentType()
                    {
                        ContentLocation = @"http://www.contoso.com/xyz.abc",
                        Name = "ItemName",

                        // Ensure content id is unique.
                        ContentId = Guid.NewGuid().ToString(),
                        ContentType = "image/jpeg",
                    };

                    switch (attachmentsType[attachmentIndex])
                    {
                        case AttachmentTypeValue.ItemAttachment:
                            itemAttachment.Item = new ItemType();
                            break;

                        case AttachmentTypeValue.MessageAttachment:
                            itemAttachment.Item = new MessageType();
                            break;

                        case AttachmentTypeValue.CalendarAttachment:
                            itemAttachment.Item = new CalendarItemType()
                            {
                                StartSpecified = true,
                                EndSpecified = true
                            };
                            break;

                        case AttachmentTypeValue.ContactAttachment:
                            itemAttachment.Item = new ContactItemType();
                            break;

                        case AttachmentTypeValue.PostAttachment:
                            itemAttachment.Item = new PostItemType();
                            break;

                        case AttachmentTypeValue.TaskAttachment:
                            itemAttachment.Item = new TaskType();
                            break;
                    }

                    attachment = itemAttachment;
                }

                attachments[attachmentIndex] = attachment;
            }

            CreateAttachmentType createAttachmentRequest = new CreateAttachmentType()
            {
                ParentItemId = new ItemIdType()
                {
                    Id = parentItemId
                },
                Attachments = attachments
            };

            return this.ATTAdapter.CreateAttachment(createAttachmentRequest);
        }

        /// <summary>
        /// Gets an attachment from an item.
        /// </summary>
        /// <param name="bodyType">Represents the format of the body text in a response.</param>
        /// <param name="includeMimeContent">Indicates whether the MIME content of an item or attachment is returned in a response. </param>
        /// <param name="attachmentIds">Contains the identifiers of the attachments to return in the response.</param>
        /// <returns>A response message for "GetAttachment" operation.</returns>
        protected GetAttachmentResponseType CallGetAttachmentOperation(BodyTypeResponseType bodyType, bool includeMimeContent, params AttachmentIdType[] attachmentIds)
        {
            GetAttachmentType getAttachmentRequest = new GetAttachmentType()
             {
                 AttachmentIds = attachmentIds,

                 AttachmentShape = new AttachmentResponseShapeType()
                 {
                     BodyType = bodyType,
                     BodyTypeSpecified = true,
                     IncludeMimeContent = includeMimeContent,
                     IncludeMimeContentSpecified = true,
                     AdditionalProperties = new BasePathToElementType[]
                    {
                        new PathToIndexedFieldType()
                        {
                             FieldURI = DictionaryURIType.itemInternetMessageHeader,
                             FieldIndex = string.Empty
                        },
                        new PathToUnindexedFieldType()
                        {
                            FieldURI = UnindexedFieldURIType.itemDateTimeCreated
                        }
                    }
                 }
             };

            return this.ATTAdapter.GetAttachment(getAttachmentRequest);
        }

        /// <summary>
        /// Deletes an attachment from an item.
        /// </summary>
        /// <param name="attachmentIds">Contains the identifiers of the attachments to be deleted.</param>
        /// <returns>A response message for "DeleteAttachment" operation.</returns>
        protected DeleteAttachmentResponseType CallDeleteAttachmentOperation(params AttachmentIdType[] attachmentIds)
        {
            DeleteAttachmentType deleteAttachmentRequest = new DeleteAttachmentType()
            {
                AttachmentIds = attachmentIds
            };

            return this.ATTAdapter.DeleteAttachment(deleteAttachmentRequest);
        }

        /// <summary>
        /// Configure the SOAP header before calling operations.
        /// </summary>
        protected void ConfigureSOAPHeader()
        {
            // Configure mailbox culture value.
            MailboxCultureType mailboxCulture = new MailboxCultureType();
            string culture = Common.GetConfigurationPropertyValue("MailboxCulture", this.Site);
            mailboxCulture.Value = culture;

            // Set the value of ExchangeImpersonation.
            ExchangeImpersonationType impersonation = new ExchangeImpersonationType();
            impersonation.ConnectingSID = new ConnectingSIDType();
            impersonation.ConnectingSID.Item = Common.GetConfigurationPropertyValue("UserName", this.Site) + "@" + Common.GetConfigurationPropertyValue("Domain", this.Site);

            // Configure time zone value.
            TimeZoneDefinitionType timezoneDefin = new TimeZoneDefinitionType();
            timezoneDefin.Id = "Eastern Standard Time";
            TimeZoneContextType timezoneContext = new TimeZoneContextType();
            timezoneContext.TimeZoneDefinition = timezoneDefin;

            Dictionary<string, object> headerValues = new Dictionary<string, object>();
            headerValues.Add("MailboxCulture", mailboxCulture);
            headerValues.Add("ExchangeImpersonation", impersonation);
            headerValues.Add("TimeZoneContext", timezoneContext);
            this.ATTAdapter.ConfigureSOAPHeader(headerValues);
        }

        /// <summary>
        /// Create a message in inbox .
        /// </summary>
        /// <returns>Id of the created message.</returns>
        protected ItemIdType CreateMessage()
        {
            CreateItemType createItemRequest = new CreateItemType()
            {
                MessageDisposition = MessageDispositionType.SaveOnly,
                MessageDispositionSpecified = true,
                SavedItemFolderId = new TargetFolderIdType()
                {
                    Item = new DistinguishedFolderIdType()
                    {
                        Id = DistinguishedFolderIdNameType.inbox
                    }
                },

                Items = new NonEmptyArrayOfAllItemsType()
                {
                    Items = new ItemType[]
                    {
                        new MessageType()
                        {     
                             Subject = Common.GenerateResourceName(this.Site, "Attachment parent message  ")
                        }
                    }
                }
            };

            CreateItemResponseType createItemResponse = this.COREAdapter.CreateItem(createItemRequest);
            ItemInfoResponseMessageType itemInfo = createItemResponse.ResponseMessages.Items[0] as ItemInfoResponseMessageType;
            MessageType message = (MessageType)itemInfo.Items.Items[0];
            return message.ItemId;
        }

        /// <summary>
        /// Delete a specific message.
        /// </summary>
        /// <param name="messageId">The Id of the message to be deleted.</param>
        /// <returns>True if the delete operation success, otherwise false.</returns>
        protected bool DeleteMessage(string messageId)
        {
            DeleteItemType deleteItemRequest = new DeleteItemType()
            {
                DeleteType = DisposalType.HardDelete,
                ItemIds = new BaseItemIdType[] 
                {
                    new ItemIdType()
                    {
                        Id = messageId
                    }
                }
            };

            DeleteItemResponseType deleteItemResponse = this.COREAdapter.DeleteItem(deleteItemRequest);
            ResponseMessageType responseMessage = deleteItemResponse.ResponseMessages.Items[0] as ResponseMessageType;
            return responseMessage.ResponseClass == ResponseClassType.Success;
        }
        #endregion
    }
}
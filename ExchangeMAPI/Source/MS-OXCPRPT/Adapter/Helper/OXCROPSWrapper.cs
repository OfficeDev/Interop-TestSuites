namespace Microsoft.Protocols.TestSuites.MS_OXCPRPT
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// MS_OXCPRPTAdapter protocol adapter.
    /// </summary>
    public partial class MS_OXCPRPTAdapter : ManagedAdapterBase, IMS_OXCPRPTAdapter
    {
       #region Const values
        /// <summary>
        /// The logon id.
        /// </summary>
        private const byte LogonId = 0;

        /// <summary>
        /// Reserved value.
        /// </summary>
        private const byte ReservedValue = 0;
        #endregion 

        #region Variables

        /// <summary>
        /// Raw data value.
        /// </summary>
        private byte[] rawDataValue;

        /// <summary>
        /// Response value.
        /// </summary>
        private IDeserializable responseValue;

        /// <summary>
        /// Response SOH value.
        /// </summary>
        private List<List<uint>> responseSOHsValue;
        #endregion

        /// <summary>
        /// This ROP logs on to a private mailbox or public folder. 
        /// </summary>
        /// <param name="logonType">This type specifies ongoing action on the private mailbox or public folder.</param>
        /// <param name="logonResponse">The response of this ROP.</param>
        /// <param name="userDN">This string specifies which mailbox to log on to.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The handle of logon object.</returns>
        private uint RopLogon(LogonType logonType, out RopLogonResponse logonResponse, string userDN, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;
            userDN += "\0";
            uint insideObjHandle = 0;

            RopLogonRequest logonRequest = new RopLogonRequest()
            {
                RopId = (byte)RopId.RopLogon,
                LogonId = LogonId,
                OutputHandleIndex = (byte)HandleIndex.FirstIndex,
                StoreState = (uint)StoreState.None,

                // Set parameters for public folder logon type.
                LogonFlags = logonType == LogonType.PublicFolder ? (byte)LogonFlags.PublicFolder : (byte)LogonFlags.Private,
                OpenFlags = logonType == LogonType.PublicFolder ? (uint)(OpenFlags.UsePerMDBReplipMapping | OpenFlags.Public) : (uint)OpenFlags.UsePerMDBReplipMapping,

                // Set EssdnSize to 0, which specifies the size of the ESSDN field.
                EssdnSize = logonType == LogonType.PublicFolder ? (ushort)0 : (ushort)Encoding.ASCII.GetByteCount(userDN),
                Essdn = logonType == LogonType.PublicFolder ? null : Encoding.ASCII.GetBytes(userDN),
            };

            this.responseSOHsValue = this.ProcessSingleRop(logonRequest, insideObjHandle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            logonResponse = (RopLogonResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, logonResponse.ReturnValue, string.Format("Logon Failed! Error: 0x{0:X8}", logonResponse.ReturnValue));
            }

            return this.responseSOHsValue[0][logonResponse.OutputHandleIndex];
        }

        #region FolderROPs
        /// <summary>
        /// This ROP creates a new subfolder. 
        /// </summary>
        /// <param name="handle">The handle to operate.</param>
        /// <param name="createFolderResponse">The response of this ROP.</param>
        /// <param name="displayName">The name of the created folder. </param>
        /// <param name="comment">The folder comment that is associated with the created folder.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The handle of new folder.</returns>
        private uint RopCreateFolder(uint handle, out RopCreateFolderResponse createFolderResponse, string displayName, string comment, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopCreateFolderRequest createFolderRequest = new RopCreateFolderRequest()
            {
                RopId = (byte)RopId.RopCreateFolder,
                LogonId = LogonId,
                InputHandleIndex = (byte)HandleIndex.FirstIndex,
                OutputHandleIndex = (byte)HandleIndex.SecondIndex,
                FolderType = (byte)FolderType.Genericfolder,
                UseUnicodeStrings = Convert.ToByte(false),
                OpenExisting = Convert.ToByte(true),
                Reserved = ReservedValue,
                DisplayName = Encoding.ASCII.GetBytes(displayName + "\0"),
                Comment = Encoding.ASCII.GetBytes(comment + "\0")
            };

            this.responseSOHsValue = this.ProcessSingleRop(createFolderRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            createFolderResponse = (RopCreateFolderResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, createFolderResponse.ReturnValue, string.Format("RopCreateFolder Failed! Error: 0x{0:X8}", createFolderResponse.ReturnValue));
            }

            return this.responseSOHsValue[0][createFolderResponse.OutputHandleIndex];
        }

        /// <summary>
        /// This ROP deletes all messages and subfolders from a folder. 
        /// </summary>
        /// <param name="handle">The handle to operate.</param>
        /// <param name="wantAsynchronous">This value specifies whether the operation is to be executed asynchronously with status reported via RopProgress.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopEmptyFolderResponse RopEmptyFolder(uint handle, byte wantAsynchronous, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopEmptyFolderRequest emptyFolderRequest = new RopEmptyFolderRequest()
            {
                RopId = (byte)RopId.RopEmptyFolder,
                LogonId = LogonId,
                InputHandleIndex = (byte)HandleIndex.FirstIndex,
                WantAsynchronous = wantAsynchronous,
                WantDeleteAssociated = Convert.ToByte(true),
            };

            this.responseSOHsValue = this.ProcessSingleRop(emptyFolderRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            RopEmptyFolderResponse emptyFolderResponse = (RopEmptyFolderResponse)this.responseValue;
            if (needVerify)
            {
                Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, emptyFolderResponse.ReturnValue, string.Format("RopEmptyFolder Failed! Error: 0x{0:X8}", emptyFolderResponse.ReturnValue));
            }

            return emptyFolderResponse;
        }

        /// <summary>
        /// This ROP deletes a folder and its contents.
        /// </summary>
        /// <param name="handle">The handle of the folder to be deleted.</param>
        /// <param name="folderId">The Id of the folder to be deleted.</param>
        /// <returns>The response of this ROP</returns>
        private RopDeleteFolderResponse RopDeleteFolder(uint handle, ulong folderId)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopDeleteFolderRequest deleteFolderRequest = new RopDeleteFolderRequest()
            {
                RopId = (byte)RopId.RopDeleteFolder,
                LogonId = LogonId,
                InputHandleIndex = (byte)HandleIndex.FirstIndex,
                DeleteFolderFlags = 0x15,
                FolderId = folderId,
            };

            this.responseSOHsValue = this.ProcessSingleRop(deleteFolderRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            RopDeleteFolderResponse deleteFolderResponse = (RopDeleteFolderResponse)this.responseValue;
            this.Site.Assert.IsNotNull(deleteFolderResponse, string.Format("RopDeleteFolderResponse Failed! Error: 0x{0:X8}", deleteFolderResponse.ReturnValue));
            return deleteFolderResponse;
        }

        /// <summary>
        /// This ROP opens an existing folder in a mailbox.  
        /// </summary>
        /// <param name="handle">The handle to operate</param>
        /// <param name="openFolderResponse">The response of this ROP</param>
        /// <param name="folderId">The identifier of the folder to be opened.</param>
        /// <param name="needVerify">Whether need to verify the response</param>
        /// <returns>The handle of the opened folder</returns>
        private uint RopOpenFolder(uint handle, out RopOpenFolderResponse openFolderResponse, ulong folderId, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopOpenFolderRequest openFolderRequest = new RopOpenFolderRequest()
            {
                RopId = (byte)RopId.RopOpenFolder,
                LogonId = LogonId,
                InputHandleIndex = (byte)HandleIndex.FirstIndex,
                OutputHandleIndex = (byte)HandleIndex.SecondIndex,
                FolderId = folderId,

                // Open an existing folder with None value for OpenModeFlags flag.
                OpenModeFlags = (byte)FolderOpenModeFlags.None
            };

            this.responseSOHsValue = this.ProcessSingleRop(openFolderRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            openFolderResponse = (RopOpenFolderResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, openFolderResponse.ReturnValue, string.Format("RopOpenFolderResponse Failed! Error: 0x{0:X8}", openFolderResponse.ReturnValue));
            }

            return this.responseSOHsValue[0][openFolderResponse.OutputHandleIndex];
        }

        #endregion

        #region MessageROPs
        /// <summary>
        /// This ROP creates a Message object in a mailbox. 
        /// </summary>
        /// <param name="handle">The handle to operate.</param>
        /// <param name="folderId">This value identifies the parent folder.</param>
        /// <param name="associatedFlag">This flag specifies whether the message is a folder associated information (FAI) message.</param>
        /// <param name="createMessageResponse">The response of this ROP.</param>
        /// <param name="needVerify">Whether need to verify the response</param>
        /// <returns>The handle of the created message.</returns>
        private uint RopCreateMessage(uint handle, ulong folderId, byte associatedFlag, out RopCreateMessageResponse createMessageResponse, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopCreateMessageRequest createMessageRequest = new RopCreateMessageRequest()
            {
                RopId = (byte)RopId.RopCreateMessage,
                LogonId = LogonId,
                InputHandleIndex = (byte)HandleIndex.FirstIndex,
                OutputHandleIndex = (byte)HandleIndex.SecondIndex,
                CodePageId = (ushort)CodePageId.SameAsLogonObject,
                FolderId = folderId,
                AssociatedFlag = associatedFlag
            };

            this.responseSOHsValue = this.ProcessSingleRop(createMessageRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            createMessageResponse = (RopCreateMessageResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, createMessageResponse.ReturnValue, string.Format("RopCreateMessageResponse Failed! Error: 0x{0:X8}", createMessageResponse.ReturnValue));
            }

            return this.responseSOHsValue[0][createMessageResponse.OutputHandleIndex];
        }

        /// <summary>
        /// This ROP opens an existing message in a mailbox. 
        /// </summary>
        /// <param name="handle">The handle to operate</param>
        /// <param name="folderId">The parent folder of the message to be opened.</param>
        /// <param name="messageId">The identifier of the message to be opened.</param>
        /// <param name="openMessageResponse">The response of this ROP.</param>
        /// <param name="needVerify">Whether need to verify the response</param>
        /// <returns>The handle of the opened message</returns>
        private uint RopOpenMessage(uint handle, ulong folderId, ulong messageId, out RopOpenMessageResponse openMessageResponse, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopOpenMessageRequest openMessageRequest = new RopOpenMessageRequest()
            {
                RopId = (byte)RopId.RopOpenMessage,
                LogonId = LogonId,
                InputHandleIndex = (byte)HandleIndex.FirstIndex,
                OutputHandleIndex = (byte)HandleIndex.SecondIndex,

                // Set CodePageId to SameAsLogonObject, which indicates it uses the same code page as the one for logon object.
                CodePageId = (ushort)CodePageId.SameAsLogonObject,
                FolderId = folderId,

                // Set OpenModeFlags to read and write for further operation.
                OpenModeFlags = (byte)MessageOpenModeFlags.ReadWrite,
                MessageId = messageId
            };

            this.responseSOHsValue = this.ProcessSingleRop(openMessageRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            openMessageResponse = (RopOpenMessageResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, openMessageResponse.ReturnValue, string.Format("RopOpenMessageResponse Failed! Error: 0x{0:X8}", openMessageResponse.ReturnValue));
            }

            return this.responseSOHsValue[0][openMessageResponse.OutputHandleIndex];
        }

        /// <summary>
        /// This ROP creates a new attachment on a message. 
        /// </summary>
        /// <param name="handle">The handle to operate</param>
        /// <param name="createAttachmentResponse">The response of this ROP.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The index of the output handle for the response</returns>
        private uint RopCreateAttachment(uint handle, out RopCreateAttachmentResponse createAttachmentResponse, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopCreateAttachmentRequest createAttachmentRequest = new RopCreateAttachmentRequest()
            {
                RopId = (byte)RopId.RopCreateAttachment,
                LogonId = LogonId,
                InputHandleIndex = (byte)HandleIndex.FirstIndex,
                OutputHandleIndex = (byte)HandleIndex.SecondIndex,
            };

            this.responseSOHsValue = this.ProcessSingleRop(createAttachmentRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            createAttachmentResponse = (RopCreateAttachmentResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, createAttachmentResponse.ReturnValue, string.Format("RopCreateAttachmentResponse Failed! Error: 0x{0:X8}", createAttachmentResponse.ReturnValue));
            }

            return this.responseSOHsValue[0][createAttachmentResponse.OutputHandleIndex];
        }

        /// <summary>
        /// This ROP opens an attachment of a message. 
        /// </summary>
        /// <param name="handle">The handle to operate.</param>
        /// <param name="attachmentId">The identifier of the attachment to be opened. </param>
        /// <param name="openAttachmentResponse">The response of this ROP.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The index of the output handle for the response.</returns>
        private uint RopOpenAttachment(uint handle, uint attachmentId, out RopOpenAttachmentResponse openAttachmentResponse, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopOpenAttachmentRequest openAttachmentRequest = new RopOpenAttachmentRequest()
            {
                RopId = (byte)RopId.RopOpenAttachment,
                LogonId = LogonId,
                InputHandleIndex = (byte)HandleIndex.FirstIndex,
                OutputHandleIndex = (byte)HandleIndex.SecondIndex,
                OpenAttachmentFlags = (byte)OpenAttachmentFlags.ReadWrite,
                AttachmentID = attachmentId
            };

            this.responseSOHsValue = this.ProcessSingleRop(openAttachmentRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            openAttachmentResponse = (RopOpenAttachmentResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, openAttachmentResponse.ReturnValue, string.Format("RopOpenAttachment Failed! Error: 0x{0:X8}", openAttachmentResponse.ReturnValue));
            }

            return this.responseSOHsValue[0][openAttachmentResponse.OutputHandleIndex];
        }

        /// <summary>
        /// This ROP commits the changes made to an attachment. 
        /// </summary>
        /// <param name="handle">The handle to operate.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopSaveChangesAttachmentResponse RopSaveChangesAttachment(uint handle, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopSaveChangesAttachmentRequest saveChangesAttachmentRequest = new RopSaveChangesAttachmentRequest()
            {
                RopId = (byte)RopId.RopSaveChangesAttachment,
                LogonId = LogonId,
                InputHandleIndex = (byte)HandleIndex.FirstIndex,
                ResponseHandleIndex = (byte)HandleIndex.SecondIndex,

                // Set the SaveFlags flag to ForceSave, which indicates the client requests server to commit the changes. 
                SaveFlags = (byte)SaveFlags.ForceSave
            };

            this.responseSOHsValue = this.ProcessSingleRop(saveChangesAttachmentRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            RopSaveChangesAttachmentResponse saveChangesAttachmentReponse = (RopSaveChangesAttachmentResponse)this.responseValue;
            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, saveChangesAttachmentReponse.ReturnValue, string.Format("RopSaveChangesAttachment Failed! Error: 0x{0:X8}", saveChangesAttachmentReponse.ReturnValue));
            }

            return saveChangesAttachmentReponse;
        }

        /// <summary>
        /// This ROP opens an attachment as a message. 
        /// </summary>
        /// <param name="handle">The handle to operate.</param>
        /// <param name="openEmbeddedMessageResponse">The response of this ROP.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The index of the output handle of the response.</returns>
        private uint RopOpenEmbeddedMessage(uint handle, out RopOpenEmbeddedMessageResponse openEmbeddedMessageResponse, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopOpenEmbeddedMessageRequest openEmbeddedMessageRequest = new RopOpenEmbeddedMessageRequest()
            {
                RopId = (byte)RopId.RopOpenEmbeddedMessage,

                // The logonId 0x00 is associated with RopOpenEmbeddedMessage.
                LogonId = 0x00,

                // This index specifies the location 0x00 in the Server Object Handle Table where the handle for the input Server Object is stored. 
                InputHandleIndex = (byte)HandleIndex.FirstIndex,

                // This index specifies the location 0x01 in the Server Object Handle Table where the handle for the output Server Object is stored. 
                OutputHandleIndex = (byte)HandleIndex.SecondIndex,
                CodePageId = 0x0FFF,
                OpenModeFlags = 0x02
            };

            this.responseSOHsValue = this.ProcessSingleRop(openEmbeddedMessageRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            openEmbeddedMessageResponse = (RopOpenEmbeddedMessageResponse)this.responseValue;

            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, openEmbeddedMessageResponse.ReturnValue, string.Format("RopOpenEmbeddedMessage Failed! Error: 0x{0:X8}", openEmbeddedMessageResponse.ReturnValue));
            }

            return this.responseSOHsValue[0][openEmbeddedMessageResponse.OutputHandleIndex];
        }

        /// <summary>
        /// This ROP commits the changes made to a message. 
        /// </summary>
        /// <param name="handle">The handle to operate.</param>
        /// <param name="needVerify">Whether need to verify the response.</param>
        /// <returns>The response of this ROP.</returns>
        private RopSaveChangesMessageResponse RopSaveChangesMessage(uint handle, bool needVerify)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopSaveChangesMessageRequest saveChangesMessageRequest = new RopSaveChangesMessageRequest()
            {
                RopId = (byte)RopId.RopSaveChangesMessage,
                LogonId = LogonId,
                InputHandleIndex = (byte)HandleIndex.FirstIndex,
                ResponseHandleIndex = (byte)HandleIndex.SecondIndex,

                // Set the SaveFlags flag to ForceSave, which indicates the client requests server to commit the changes. 
                SaveFlags = (byte)SaveFlags.ForceSave,
            };

            this.responseSOHsValue = this.ProcessSingleRop(saveChangesMessageRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
            RopSaveChangesMessageResponse saveChangesMessageResponse = (RopSaveChangesMessageResponse)this.responseValue;

            if (needVerify)
            {
                this.Site.Assert.AreEqual((uint)RopResponseType.SuccessResponse, saveChangesMessageResponse.ReturnValue, string.Format("RopSaveChangesMessage Failed! Error: 0x{0:X8}", saveChangesMessageResponse.ReturnValue));
            }

            return saveChangesMessageResponse;
        }
        #endregion

        /// <summary>
        /// This ROP releases all resources associated with a server object. 
        /// </summary>
        /// <param name="handle">The handle to operate.</param>
        private void RopRelease(uint handle)
        {
            this.rawDataValue = null;
            this.responseValue = null;
            this.responseSOHsValue = null;

            RopReleaseRequest releaseRequest = new RopReleaseRequest()
            {
                RopId = (byte)RopId.RopRelease,
                LogonId = LogonId,
                InputHandleIndex = (byte)HandleIndex.FirstIndex
            };

            // The RopRelease ROP doesn't return response from server.
            this.responseSOHsValue = this.ProcessSingleRop(releaseRequest, handle, ref this.responseValue, ref this.rawDataValue, RopResponseType.SuccessResponse);
        }
    }
}
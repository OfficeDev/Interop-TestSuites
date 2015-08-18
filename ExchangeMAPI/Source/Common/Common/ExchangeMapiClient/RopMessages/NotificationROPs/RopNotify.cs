namespace Microsoft.Protocols.TestSuites.Common
{
    using System;
    using System.Runtime.InteropServices;

    /// <summary>
    /// RopNotify response buffer structure.
    /// </summary>
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct RopNotifyResponse : IDeserializable
    {
        /// <summary>
        /// This value specifies the Type of remote operation. For this operation, this field is set to 0x2A.
        /// </summary>
        public byte RopId;

        /// <summary>
        /// This handle specifies the notification server object associated with this notification event.
        /// </summary>
        public uint NotificationHandle;

        /// <summary>
        /// This value specifies the logon associated with this notification event.
        /// </summary>
        public byte LogonId;

        /// <summary>
        /// Various structures. The notification structures that can be found here are specified in [MS-OXCDATA].
        /// </summary>
        public byte[] NotificationData;

        /// <summary>
        /// Gets or sets the NotificationFlags.This value specifies the type of the notification and availability of the notification data fields
        /// </summary>
        public ushort NotificationFlags { get; set; }

        /// <summary>
        /// Gets the Notification Type.
        /// </summary>
        public NotificationType NotificationType
        {
            get
            {
                return (NotificationType)(this.NotificationFlags & 0x0FFF);
            }
        }

        /// <summary>
        /// Gets the FriendlyTableEvent. For debug use
        /// </summary>
        public string FriendlyTableEvent
        {
            get
            {
                return NotificationType.ToString() + (this.TableEvent == EventTypeOfTable.NONE ? null : ("  " + this.TableEvent.ToString()));
            }
        }

        /// <summary>
        /// Gets or sets the TableEventType.This value specifies the Subtype of the notification for a TableModified event.
        /// </summary>
        public ushort? TableEventType { get; set; }

        /// <summary>
        /// Gets the Table Event.
        /// </summary>
        public EventTypeOfTable TableEvent
        {
            get
            {
                return (EventTypeOfTable)this.TableEventType;
            }
        }

        /// <summary>
        /// Gets or sets the value.This value specifies Folder ID of the item that is triggering this notification.
        /// </summary>
        public ulong? TableRowFolderID { get; set; }

        /// <summary>
        /// Gets or sets the value.This value specifies Message ID of the item triggering this notification. 
        /// </summary>
        public ulong? TableRowMessageID { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies An identifier of the instance of the previous row in the table. 
        /// </summary>
        public uint? TableRowInstance { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Old folder ID of the item triggering this notification. 
        /// </summary>
        public ulong? InsertAfterTableRowFolderID { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Old message ID of the item triggering this notification. 
        /// </summary>
        public ulong? InsertAfterTableRowID { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies An identifier of the instance of the old row in the table.
        /// </summary>
        public uint? InsertAfterTableRowInstance { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Length of table row data. 
        /// </summary>
        public ushort? TableRowDataSize { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Table row data. 
        /// </summary>
        public byte[] TableRowData { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies that Set to TRUE (non-zero) if folder hierarchy has changed. Set to FALSE (zero) otherwise. 
        /// </summary>
        public byte? HierarchyChanged { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Number of folder IDs. 
        /// </summary>
        public uint? FolderIDNumber { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Folder IDs. 
        /// </summary>
        public GlobalIdentifier[] FolderIDs { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Folder CNs. 
        /// </summary>
        public uint[] ICSChangeNumbers { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Folder ID of the item triggering the event. 
        /// </summary>
        public ulong? FolderId { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Message ID of the item triggering the event. 
        /// </summary>
        public ulong? MessageId { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Folder ID of the parent folder of the item triggering the event.
        /// </summary>
        public ulong? ParentFolderId { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Old folder ID of the item triggering the event.
        /// </summary>
        public ulong? OldFolderId { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Old message ID of the item triggering the event.
        /// </summary>
        public ulong? OldMessageId { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Old parent folder ID of the item triggering the event.
        /// </summary>
        public ulong? OldParentFolderId { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Number of property tags.
        /// </summary>
        public ushort? TagCount { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies List of IDs of properties that have changed.
        /// </summary>
        public uint[] Tags { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Total number of items in a folder triggering this event. 
        /// </summary>
        public uint? TotalMessageCount { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Number of unread items in a folder triggering this event. 
        /// </summary>
        public uint? UnreadMessageCount { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies Message flags of new mail that has been received. 
        /// </summary>
        public uint? MessageFlags { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies that Set to TRUE (non-zero) if MessageClass is in Unicode. 
        /// </summary>
        public byte? UnicodeFlag { get; set; }

        /// <summary>
        /// Gets or sets the value. This value specifies that Null-terminated string containing the message class of the new mail.
        /// </summary>
        public byte[] MessageClass { get; set; }

        /// <summary>
        /// Gets or sets the value. This struct specifies the fields of RopNotify is available or not.
        /// </summary>
        public AvailableFieldsType AvailableFields { get; set; }

        /// <summary>
        /// Deserialize the ROP response buffer.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of response buffer struct.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            this.AvailableFields = new AvailableFieldsType();
            int index = startIndex;
            this.RopId = ropBytes[index++];
            this.NotificationHandle = (uint)BitConverter.ToInt32(ropBytes, index);
            index += 4;
            this.LogonId = ropBytes[index++];
            this.NotificationData = new byte[ropBytes.Length - index];

            this.NotificationFlags = (ushort)BitConverter.ToUInt16(ropBytes, index);
            index += 2;
            if ((this.NotificationFlags & 0x0100) == 0x0100)
            {
                this.TableEventType = (ushort)BitConverter.ToUInt16(ropBytes, index);
                index += 2;
            }

            if (this.HasTableRowFolderId())
            {
                this.TableRowFolderID = (ulong)BitConverter.ToUInt64(ropBytes, index);
                this.AvailableFields.IsTableRowFolderIDAvailable = true;
                index += 8;
                if ((this.NotificationFlags & 0x8000) == 0x8000)
                {
                    this.TableRowMessageID = (ulong)BitConverter.ToUInt64(ropBytes, index);
                    this.AvailableFields.IsTableRowMessageIDAvailable = true;
                    index += 8;
                    this.TableRowInstance = (uint)BitConverter.ToUInt32(ropBytes, index);
                    this.AvailableFields.IsTableRowPreviousInstanceAvailable = true;
                    index += 4;
                }

                if (this.TableEventType == 0x03 || this.TableEventType == 0x05)
                {
                    this.InsertAfterTableRowFolderID = (ulong)BitConverter.ToUInt64(ropBytes, index);
                    this.AvailableFields.IsTableRowOldFolderIDAvailable = true;
                    index += 8;
                    if ((this.NotificationFlags & 0x8000) == 0x8000)
                    {
                        this.InsertAfterTableRowID = (ulong)BitConverter.ToUInt64(ropBytes, index);
                        this.AvailableFields.IsTableRowOldMessageIDAvailable = true;
                        index += 8;

                        this.InsertAfterTableRowInstance = (uint)BitConverter.ToUInt32(ropBytes, index);
                        index += 4;
                    }

                    this.TableRowDataSize = (ushort)BitConverter.ToUInt16(ropBytes, index);
                    this.AvailableFields.IsTableRowDataSizeAvailable = true;
                    index += 2;
                    this.TableRowData = new byte[this.TableRowDataSize.Value];
                    Array.Copy(ropBytes, index, this.TableRowData, 0, this.TableRowDataSize.Value);
                    index += this.TableRowDataSize.Value;
                }
            }

            if ((this.NotificationFlags & 0x0200) == 0x0200)
            {
                this.HierarchyChanged = ropBytes[index++];
                this.FolderIDNumber = (uint)BitConverter.ToUInt32(ropBytes, index);
                this.AvailableFields.IsFolderIDNumberAvailable = true;
                index += 4;
                this.FolderIDs = new GlobalIdentifier[this.FolderIDNumber.Value];
                this.ICSChangeNumbers = new uint[this.FolderIDNumber.Value];

                for (int i = 0; i < this.FolderIDNumber; i++)
                {
                    index += this.FolderIDs[i].Deserialize(ropBytes, index);
                }

                for (int i = 0; i < this.FolderIDNumber; i++)
                {
                    this.ICSChangeNumbers[i] = (uint)BitConverter.ToUInt32(ropBytes, index);
                    index += 4;
                }
            }
            else
            {
                // when the field HierarchyChanged is not available,set value 0xFF to it.
                this.HierarchyChanged = 0xFF;
            }

            if (this.HasFolderId())
            {
                this.FolderId = (ulong)BitConverter.ToUInt64(ropBytes, index);
                index += 8;
                if ((this.NotificationFlags & 0x8000) == 0x8000)
                {
                    this.MessageId = (ulong)BitConverter.ToUInt64(ropBytes, index);
                    index += 8;
                }
            }

            if (this.HasParentFolderId())
            {
                this.ParentFolderId = (ulong)BitConverter.ToUInt64(ropBytes, index);
                index += 8;
            }

            if ((this.NotificationFlags & 0x0020) == 0x0020 ||
                (this.NotificationFlags & 0x0040) == 0x0040)
            {
                this.OldFolderId = (ulong)BitConverter.ToUInt64(ropBytes, index);
                index += 8;
                if ((this.NotificationFlags & 0x8000) == 0x8000)
                {
                    this.OldMessageId = (ulong)BitConverter.ToUInt64(ropBytes, index);
                    index += 8;
                }
                else
                {
                    this.OldParentFolderId = (ulong)BitConverter.ToUInt64(ropBytes, index);
                    index += 8;
                }
            }

            if ((this.NotificationFlags & 0x0004) == 0x0004 ||
                (this.NotificationFlags & 0x0010) == 0x0010)
            {
                this.TagCount = (ushort)BitConverter.ToUInt16(ropBytes, index);
                index += 2;
            }

            if (this.TagCount > 0 && this.TagCount != 0xFFFF)
            {
                this.Tags = new uint[this.TagCount.Value];
                for (int i = 0; i < this.TagCount; i++)
                {
                    this.Tags[i] = (uint)BitConverter.ToUInt32(ropBytes, index);
                    index += 4;
                }
            }

            if ((this.NotificationFlags & 0x1000) == 0x1000)
            {
                this.TotalMessageCount = (uint)BitConverter.ToUInt32(ropBytes, index);
                this.AvailableFields.IsTotalMessageCountAvailable = true;
                index += 4;
            }

            if ((this.NotificationFlags & 0x2000) == 0x2000)
            {
                this.UnreadMessageCount = (uint)BitConverter.ToUInt32(ropBytes, index);
                this.AvailableFields.IsUnreadMessageCountAvailable = true;
                index += 4;
            }

            if ((this.NotificationFlags & 0x0fff) == 0x0002)
            {
                this.MessageFlags = (uint)BitConverter.ToUInt32(ropBytes, index);
                index += 4;
                this.UnicodeFlag = ropBytes[index++];
                this.ParseString(ref index, ref ropBytes);
            }

            return index - startIndex;
        }

        #region Common method for verify requirments of RopNotify response
        /// <summary>
        /// Check TableEventType value is avaliable value.
        /// </summary>
        /// <returns>True or false</returns>
        public bool IsAvailableTableEventType()
        {
            bool isAvailableTableEventType =
                this.TableEventType == (ushort)EventTypeOfTable.TableChanged ||
                this.TableEventType == (ushort)EventTypeOfTable.TableRestrictionChanged ||
                this.TableEventType == (ushort)EventTypeOfTable.TableRowAdded ||
                this.TableEventType == (ushort)EventTypeOfTable.TableRowDeleted ||
                this.TableEventType == (ushort)EventTypeOfTable.TableRowModified;

            return isAvailableTableEventType;
        }

        /// <summary>
        /// Check the TableEventType field is available and is equal to TableRowAdded(0x03), TableRowDeleted(0x04), or TableRowModified(0x05).
        /// </summary>
        /// <returns>True or false</returns>
        public bool IsTableEventTypeValueForR329()
        {
            bool isTableEventTypeForR329 =
                this.TableEventType == (ushort)EventTypeOfTable.TableRowAdded ||
                this.TableEventType == (ushort)EventTypeOfTable.TableRowDeleted ||
                this.TableEventType == (ushort)EventTypeOfTable.TableRowModified;
            return isTableEventTypeForR329;
        }

        /// <summary>
        /// Check the TableEventType field is available and is equal to TableRowAdded(0x03) or TableRowModified(0x05).
        /// </summary>
        /// <returns>True if TableEventType field is available and is equal to TableRowAdded(0x03) or TableRowModified(0x05).</returns>
        public bool IsTableEventTypeValueForR332R334R335()
        {
            bool isTableEventTypeForR332R334R335 =
                this.TableEventType == (ushort)EventTypeOfTable.TableRowAdded ||
                this.TableEventType == (ushort)EventTypeOfTable.TableRowModified;
            return isTableEventTypeForR332R334R335;
        }

        /// <summary>
        /// Check the NotificationType value in NotificationFlags is not TableModified(0x0100), StatusObjectModified(0x0200), or Reserved(0x0400).
        /// </summary>
        /// <returns>True or false</returns>
        public bool IsNotificationFlagsValueForR340()
        {
            bool isNotificationFlagsForR340 =
                (this.NotificationFlags & (ushort)NotificationType.TableModified) != (ushort)NotificationType.TableModified &&
                (this.NotificationFlags & (ushort)NotificationType.StatusObjectModified) != (ushort)NotificationType.StatusObjectModified &&
                (this.NotificationFlags & (ushort)NotificationType.Reserved) != (ushort)NotificationType.Reserved;
            return isNotificationFlagsForR340;
        }

        /// <summary>
        /// Check the NotificationType value in NotificationFlags is ObjectCreated(0x0004), ObjectDeleted(0x0008), ObjectMoved(0x0020), or ObjectCopied(0x0040) 
        /// and bit S(0x4000) and bit M(0x8000) are set or bit S(0x4000) and bit M(0x8000) are not set in NotificationFlags.
        /// </summary>
        /// <returns>True or false</returns>
        public bool IsNotificationFlagsValueForR342()
        {
            bool isNotificationFlagsForR342 =
                ((this.NotificationFlags & (ushort)NotificationType.ObjectCreated) == (ushort)NotificationType.ObjectCreated ||
                (this.NotificationFlags & (ushort)NotificationType.ObjectDeleted) == (ushort)NotificationType.ObjectDeleted ||
                (this.NotificationFlags & (ushort)NotificationType.ObjectMoved) == (ushort)NotificationType.ObjectMoved ||
                (this.NotificationFlags & (ushort)NotificationType.ObjectCopied) == (ushort)NotificationType.ObjectCopied) &&
                (((this.NotificationFlags & (ushort)FlagsBit.S) == (ushort)FlagsBit.S &&
                (this.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M) ||
                ((this.NotificationFlags & (ushort)FlagsBit.S) != (ushort)FlagsBit.S &&
                (this.NotificationFlags & (ushort)FlagsBit.M) != (ushort)FlagsBit.M));
            return isNotificationFlagsForR342;
        }

        /// <summary>
        /// Check notificationType value in NotificationFlags is ObjectMoved(0x0020) or ObjectCopied(0x0040). 
        /// </summary>
        /// <returns>True or false</returns>
        public bool IsNotificationFlagsValueForR343()
        {
            bool isNotificationFlagsForR343 =
                (this.NotificationFlags & (ushort)NotificationType.ObjectMoved) == (ushort)NotificationType.ObjectMoved ||
                (this.NotificationFlags & (ushort)NotificationType.ObjectCopied) == (ushort)NotificationType.ObjectCopied;
            return isNotificationFlagsForR343;
        }

        /// <summary>
        /// Check the NotificationType value in NotificationFlags is ObjectCreated(0x0004) or ObjectModified(0x0010).
        /// </summary>
        /// <returns>True or false</returns>
        public bool IsNotificationTypeValueForR346()
        {
            bool isNotificationTypeForR346 =
                (this.NotificationFlags & (ushort)NotificationType.ObjectCreated) == (ushort)NotificationType.ObjectCreated ||
                (this.NotificationFlags & (ushort)NotificationType.ObjectModified) == (ushort)NotificationType.ObjectModified;
            return isNotificationTypeForR346;
        }

        /// <summary>
        /// Check the NotificationType value is the member of NotificationType enumeration.
        /// </summary>
        /// <returns>True or false</returns>
        public bool IsEnumerationOfNotificationType()
        {
            bool isEnumerationOfNotificationType =
                (this.NotificationFlags & (ushort)NotificationType.NewMail) == (ushort)NotificationType.NewMail ||
                (this.NotificationFlags & (ushort)NotificationType.ObjectCreated) == (ushort)NotificationType.ObjectCreated ||
                (this.NotificationFlags & (ushort)NotificationType.ObjectDeleted) == (ushort)NotificationType.ObjectDeleted ||
                (this.NotificationFlags & (ushort)NotificationType.ObjectModified) == (ushort)NotificationType.ObjectModified ||
                (this.NotificationFlags & (ushort)NotificationType.ObjectMoved) == (ushort)NotificationType.ObjectMoved ||
                (this.NotificationFlags & (ushort)NotificationType.ObjectCopied) == (ushort)NotificationType.ObjectCopied ||
                (this.NotificationFlags & (ushort)NotificationType.SearchCompleted) == (ushort)NotificationType.SearchCompleted ||
                (this.NotificationFlags & (ushort)NotificationType.TableModified) == (ushort)NotificationType.TableModified ||
                (this.NotificationFlags & (ushort)NotificationType.StatusObjectModified) == (ushort)NotificationType.StatusObjectModified ||
                (this.NotificationFlags & (ushort)NotificationType.Reserved) == (ushort)NotificationType.Reserved;
            return isEnumerationOfNotificationType;
        }

        /// <summary>
        /// Check flags bit in NotificationFlags can be set value is:T(0x1000),U(0x2000),S(0x4000),M(0x8000) or all flags bit not be set(0x0000). 
        /// </summary>
        /// <returns>True or false</returns>
        public bool IsFlagsOfNotificationFlags()
        {
            bool isFlags =
                (this.NotificationFlags & (ushort)FlagsBit.T) == (ushort)FlagsBit.T ||
                (this.NotificationFlags & (ushort)FlagsBit.U) == (ushort)FlagsBit.U ||
                (this.NotificationFlags & (ushort)FlagsBit.S) == (ushort)FlagsBit.S ||
                (this.NotificationFlags & (ushort)FlagsBit.M) == (ushort)FlagsBit.M ||
                (this.NotificationFlags & (ushort)FlagsBit.NONE) == (ushort)FlagsBit.NONE;
            return isFlags;
        }
        #endregion

        /// <summary>
        /// Indicate if the field of ParentFolderId exist.
        /// </summary>
        /// <returns>The value indicates if the field of ParentFolderId exist</returns>
        private bool HasParentFolderId()
        {
            if (((this.NotificationFlags & 0x0004) == 0x0004 || (this.NotificationFlags & 0x0008) == 0x0008 || (this.NotificationFlags & 0x0020) == 0x0020 || (this.NotificationFlags & 0x0040) == 0x0040) &&
                (((this.NotificationFlags & 0x4000) == 0x4000 && (this.NotificationFlags & 0x8000) == 0x8000) ||
                ((this.NotificationFlags & 0x4000) != 0x4000 && (this.NotificationFlags & 0x8000) != 0x8000)))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Indicate if the field of TableRowFolderId exist.
        /// </summary>
        /// <returns>The value indicates if the field of TableRowFolderId exist</returns>
        private bool HasTableRowFolderId()
        {
            if (this.TableEventType == 0x03
                          || this.TableEventType == 0x04
                          || this.TableEventType == 0x05)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Indicate if the folder id exist
        /// </summary>
        /// <returns>True if the folder id exist</returns>
        private bool HasFolderId()
        {
            if ((this.NotificationFlags & 0x0100) != 0x0100 &&
               (this.NotificationFlags & 0x0200) != 0x0200 &&
               (this.NotificationFlags & 0x0400) != 0x0400)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Parse a string
        /// </summary>
        /// <param name="index">The index of current byte</param>
        /// <param name="ropBytes">ROPs bytes</param>
        private void ParseString(ref int index, ref byte[] ropBytes)
        {
            int strBytesLen = 0;
            bool isFound = false;

            // Unicode
            if (this.UnicodeFlag != 0x0)
            {
                for (int i = index; i < ropBytes.Length; i += 2)
                {
                    strBytesLen += 2;
                    if ((ropBytes[i] == 0) && (ropBytes[i + 1] == 0))
                    {
                        isFound = true;
                        break;
                    }
                }
            }
            else
            {
                // Find the string with '\0' end
                for (int i = index; i < ropBytes.Length; i++)
                {
                    strBytesLen++;
                    if (ropBytes[i] == 0)
                    {
                        isFound = true;
                        break;
                    }
                }
            }

            if (!isFound)
            {
                throw new ParseException("String too long or not found");
            }
            else
            {
                this.MessageClass = new byte[strBytesLen];
                Array.Copy(ropBytes, index, this.MessageClass, 0, strBytesLen - 1);
                index += strBytesLen;
            }
        }

        /// <summary>
        /// The fields TableRowDataSize,TotalMessageCount,UnreadMessageCount,FolderIDNumber,
        /// TableRowFolderID,TableRowMessageID,TableRowPreviousInstance,TableRowOldFolderID,TableRowOldMessageID
        /// </summary>
        public class AvailableFieldsType
        {
            /// <summary>
            /// Gets or sets a value indicating whether the field TableRowDataSize is available or not
            /// </summary>
            public bool IsTableRowDataSizeAvailable { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the field TotalMessageCount is available or not
            /// </summary>
            public bool IsTotalMessageCountAvailable { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the field UnreadMessageCount is available or not
            /// </summary>
            public bool IsUnreadMessageCountAvailable { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the field FolderIDNumber is available or not
            /// </summary>
            public bool IsFolderIDNumberAvailable { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the field TableRowFolderID is available or not
            /// </summary>
            public bool IsTableRowFolderIDAvailable { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the field TableRowMessageID is available or not
            /// </summary>
            public bool IsTableRowMessageIDAvailable { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the field TableRowPreviousInstance is available or not
            /// </summary>
            public bool IsTableRowPreviousInstanceAvailable { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the field TableRowOldFolderID is available or not
            /// </summary>
            public bool IsTableRowOldFolderIDAvailable { get; set; }

            /// <summary>
            /// Gets or sets a value indicating whether the field TableRowOldMessageID is available or not
            /// </summary>
            public bool IsTableRowOldMessageIDAvailable { get; set; }
        }
    }
}
namespace Microsoft.Protocols.TestSuites.Common
{
    using System;

    /// <summary>
    /// The NotificationData structure.
    /// </summary>
    public class NotificationData
    {
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
        private AvailableFieldsType AvailableFields { get; set; }

        /// <summary>
        /// Deserialize the NotificationData.
        /// </summary>
        /// <param name="ropBytes">ROPs bytes in response.</param>
        /// <param name="startIndex">The start index of this ROP.</param>
        /// <returns>The size of NotificationData struct.</returns>
        public int Deserialize(byte[] ropBytes, int startIndex)
        {
            int index = startIndex;
            this.AvailableFields = new AvailableFieldsType();

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
    }
}

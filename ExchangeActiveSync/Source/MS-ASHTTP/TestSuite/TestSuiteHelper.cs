//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_ASHTTP
{
    using System;
    using System.Collections.ObjectModel;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Request = Microsoft.Protocols.TestSuites.Common.Request;
    using Response = Microsoft.Protocols.TestSuites.Common.Response;

    /// <summary>
    /// A class contains all helper methods used in test cases.
    /// </summary>
    public static class TestSuiteHelper
    {
        /// <summary>
        /// Get status code from web exception which will be returned by IIS.
        /// </summary>
        /// <param name="webException">Web exception</param>
        /// <returns>Status code</returns>
        public static string GetStatusCodeFromException(Exception webException)
        {
            if (null == webException)
            {
                return string.Empty;
            }

            string exceptionMessage = webException.Message;
            string statusCode = string.Empty;
            if (exceptionMessage.Contains("(") && exceptionMessage.Contains(")"))
            {
                int leftParenthesis = exceptionMessage.IndexOf("(", StringComparison.OrdinalIgnoreCase);
                int rightParenthesis = exceptionMessage.IndexOf(")", StringComparison.OrdinalIgnoreCase);
                statusCode = exceptionMessage.Substring(leftParenthesis + 1, rightParenthesis - leftParenthesis - 1);
            }

            return statusCode;
        }

        /// <summary>
        /// Convert the instance of SendStringResponse to SyncResponse.
        /// </summary>
        /// <param name="syncResponseString">The SendStringResponse instance to convert.</param>
        /// <returns>The instance of SyncResponse.</returns>
        public static SyncResponse ConvertSyncResponseFromSendString(ActiveSyncResponseBase<object> syncResponseString)
        {
            SyncResponse syncResponse = new SyncResponse
            {
                ResponseDataXML = syncResponseString.ResponseDataXML,
                Headers = syncResponseString.Headers
            };

            syncResponse.DeserializeResponseData();

            return syncResponse;
        }

        /// <summary>
        /// Get policy key from Provision string response.
        /// </summary>
        /// <param name="provisionResponseString">The SendStringResponse instance of Provision command.</param>
        /// <returns>The policy key of the policy.</returns>
        public static string GetPolicyKeyFromSendString(ActiveSyncResponseBase<object> provisionResponseString)
        {
            ProvisionResponse provisionResponse = new ProvisionResponse
            {
                ResponseDataXML = provisionResponseString.ResponseDataXML
            };

            if (!string.IsNullOrEmpty(provisionResponse.ResponseDataXML))
            {
                provisionResponse.DeserializeResponseData();

                if (provisionResponse.ResponseData.Policies != null)
                {
                    Response.ProvisionPoliciesPolicy policyInResponse = provisionResponse.ResponseData.Policies.Policy;
                    if (policyInResponse != null)
                    {
                        return policyInResponse.PolicyKey;
                    }
                }
            }

            return string.Empty;
        }

        /// <summary>
        /// Get the request of Sync command.
        /// </summary>
        /// <param name="collectionId">The CollectionId of the folder to sync.</param>
        /// <param name="syncKey">The SyncKey of the latest sync.</param>
        /// <returns>The request of Sync command.</returns>
        public static SyncRequest GetSyncRequest(string collectionId, string syncKey)
        {
            // Create the Sync command request.
            Request.SyncCollection[] synCollections = new Request.SyncCollection[1];
            synCollections[0] = new Request.SyncCollection { SyncKey = syncKey, CollectionId = collectionId };
            SyncRequest syncRequest = Common.CreateSyncRequest(synCollections);
            return syncRequest;
        }

        /// <summary>
        /// Load sync response to sync store.
        /// </summary>
        /// <param name="response">The response of Sync command.</param>
        /// <returns>The sync store instance.</returns>
        public static SyncStore LoadSyncResponse(ActiveSyncResponseBase<Response.Sync> response)
        {
            if (response.ResponseData.Item == null)
            {
                return null;
            }

            SyncStore result = new SyncStore();
            Response.SyncCollectionsCollection collection = ((Response.SyncCollections)response.ResponseData.Item).Collection[0];
            for (int i = 0; i < collection.ItemsElementName.Length; i++)
            {
                switch (collection.ItemsElementName[i])
                {
                    case Response.ItemsChoiceType10.CollectionId:
                        result.CollectionId = collection.Items[i].ToString();
                        break;
                    case Response.ItemsChoiceType10.SyncKey:
                        result.SyncKey = collection.Items[i].ToString();
                        break;
                    case Response.ItemsChoiceType10.Status:
                        result.Status = Convert.ToByte(collection.Items[i]);
                        break;
                    case Response.ItemsChoiceType10.Commands:
                        Response.SyncCollectionsCollectionCommands commands = collection.Items[i] as Response.SyncCollectionsCollectionCommands;
                        if (commands != null)
                        {
                            foreach (SyncItem item in LoadAddCommands(commands))
                            {
                                result.AddCommands.Add(item);
                            }
                        }

                        break;
                    case Response.ItemsChoiceType10.Responses:
                        Response.SyncCollectionsCollectionResponses responses = collection.Items[i] as Response.SyncCollectionsCollectionResponses;
                        if (responses != null)
                        {
                            if (responses.Add != null)
                            {
                                foreach (Response.SyncCollectionsCollectionResponsesAdd add in responses.Add)
                                {
                                    result.AddResponses.Add(add);
                                }
                            }
                        }

                        break;
                }
            }

            return result;
        }

        /// <summary>
        /// Load add commands in sync response.
        /// </summary>
        /// <param name="collectionCommands">The add commands response.</param>
        /// <returns>The list of SyncItem in add commands.</returns>
        public static Collection<SyncItem> LoadAddCommands(Response.SyncCollectionsCollectionCommands collectionCommands)
        {
            if (collectionCommands.Add != null)
            {
                Collection<SyncItem> commands = new Collection<SyncItem>();
                if (collectionCommands.Add.Length > 0)
                {
                    foreach (Response.SyncCollectionsCollectionCommandsAdd addCommand in collectionCommands.Add)
                    {
                        SyncItem syncItem = new SyncItem { ServerId = addCommand.ServerId };
                        for (int i = 0; i < addCommand.ApplicationData.ItemsElementName.Length; i++)
                        {
                            switch (addCommand.ApplicationData.ItemsElementName[i])
                            {
                                case Response.ItemsChoiceType8.Subject1:
                                    syncItem.Subject = addCommand.ApplicationData.Items[i].ToString();
                                    break;
                                case Response.ItemsChoiceType8.Subject:
                                    syncItem.Subject = addCommand.ApplicationData.Items[i].ToString();
                                    break;
                            }
                        }

                        commands.Add(syncItem);
                    }
                }

                return commands;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get ServerId from Sync response.
        /// </summary>
        /// <param name="syncResponse">The response of Sync command.</param>
        /// <param name="subject">The subject of the email to get.</param>
        /// <returns>The ServerId of the email.</returns>
        public static string GetServerIdFromSyncResponse(SyncStore syncResponse, string subject)
        {
            string itemServerId = null;
            foreach (SyncItem add in syncResponse.AddCommands)
            {
                if (add.Subject == subject)
                {
                    itemServerId = add.ServerId;
                    break;
                }
            }

            return itemServerId;
        }

        /// <summary>
        /// Verify whether X-MS-RP, MS-ASProtocolCommands and MS-ASProtocolVersions headers all exist in response headers.
        /// </summary>
        /// <param name="headers">The headers returned from response.</param>
        /// <returns>Whether the X-MS-RP, MS-ASProtocolCommands and MS-ASProtocolVersions headers all exist in FolderSync response header.</returns>
        public static bool VerifySyncRequiredResponseHeaders(string[] headers)
        {
            bool headerRP = false;
            bool headerASProtocolCommands = false;
            bool headerASProtocolVersions = false;
            foreach (string header in headers)
            {
                if (header == "X-MS-RP")
                {
                    headerRP = true;
                }

                if (header == "MS-ASProtocolCommands")
                {
                    headerASProtocolCommands = true;
                }

                if (header == "MS-ASProtocolVersions")
                {
                    headerASProtocolVersions = true;
                }
            }

            return headerRP && headerASProtocolCommands && headerASProtocolVersions;
        }

        /// <summary>
        /// Get whether retry is needed when get ServerId.
        /// </summary>
        /// <param name="shouldBeGotten">Whether the item should be gotten.</param>
        /// <returns>If the retry is needed, return true, otherwise, return false.</returns>
        public static bool IsRetryNeeded(string shouldBeGotten)
        {
            if (shouldBeGotten == "T" || shouldBeGotten == "1")
            {
                return true;
            }

            return false;
        }

        /// <summary>
        /// Get the calendar content string.
        /// </summary>
        /// <param name="fromAddress">The email address of the meeting request sender.</param>
        /// <param name="toAddress">The email address of the meeting request receiver.</param>
        /// <param name="meetingRequestSubject">The subject of the meeting request.</param>
        /// <param name="occurrences">The number of occurrences of the meeting request.</param>
        /// <returns>The created meeting request mime.</returns>
        public static string CreateCalendarContent(string fromAddress, string toAddress, string meetingRequestSubject, string occurrences)
        {
            System.Text.StringBuilder icsBuilder = new System.Text.StringBuilder();
            icsBuilder.AppendLine("BEGIN:VCALENDAR");
            icsBuilder.AppendLine("PRODID:-//Microsoft Protocols TestSuites");
            icsBuilder.AppendLine("VERSION:2.0");
            icsBuilder.AppendLine("METHOD:REQUEST");

            icsBuilder.AppendLine("X-MS-OLK-FORCEINSPECTOROPEN:TRUE");
            icsBuilder.AppendLine("BEGIN:VTIMEZONE");
            icsBuilder.AppendLine("TZID:Universal Time");
            icsBuilder.AppendLine("BEGIN:STANDARD");
            icsBuilder.AppendLine("DTSTART:16011104T020000");
            icsBuilder.AppendLine("RRULE:FREQ=YEARLY;BYDAY=1SU;BYMONTH=11");
            icsBuilder.AppendLine("TZOFFSETFROM:-0000");
            icsBuilder.AppendLine("TZOFFSETTO:+0000");
            icsBuilder.AppendLine("END:STANDARD");
            icsBuilder.AppendLine("BEGIN:DAYLIGHT");
            icsBuilder.AppendLine("DTSTART:16010311T020000");
            icsBuilder.AppendLine("RRULE:FREQ=YEARLY;BYDAY=2SU;BYMONTH=3");
            icsBuilder.AppendLine("TZOFFSETFROM:-0000");
            icsBuilder.AppendLine("TZOFFSETTO:+0000");
            icsBuilder.AppendLine("END:DAYLIGHT");
            icsBuilder.AppendLine("END:VTIMEZONE");

            icsBuilder.AppendLine("BEGIN:VEVENT");
            icsBuilder.AppendLine(string.Format("DTSTART:{0:yyyyMMddTHH0000Z}", DateTime.UtcNow.AddHours(1)));
            icsBuilder.AppendLine(string.Format("DTSTAMP:{0:yyyyMMddTHHmmssZ}", DateTime.UtcNow));
            icsBuilder.AppendLine(string.Format("DTEND:{0:yyyyMMddTHH0000Z}", DateTime.UtcNow.AddHours(2)));
            icsBuilder.AppendLine(string.Format("LOCATION:{0}", "Meeting room one"));
            icsBuilder.AppendLine(string.Format("UID:{0}", Guid.NewGuid().ToString()));
            icsBuilder.AppendLine(string.Format("DESCRIPTION:Meeting Request"));
            icsBuilder.AppendLine(string.Format("X-ALT-DESC;FMTTYPE=text/html:Meeting Request"));
            icsBuilder.AppendLine(string.Format("SUMMARY:{0}", meetingRequestSubject));
            icsBuilder.AppendLine(string.Format("ORGANIZER:MAILTO:{0}", fromAddress));
            icsBuilder.AppendLine(string.Format("ATTENDEE;CN=\"{0}\";RSVP=TRUE:mailto:{1}", toAddress.Split('@')[0], toAddress));
            
            if (!string.IsNullOrEmpty(occurrences))
            {
                icsBuilder.AppendLine("RRULE:FREQ=DAILY;COUNT=" + occurrences.ToString());
            }

            icsBuilder.AppendLine("BEGIN:VALARM");
            icsBuilder.AppendLine("TRIGGER:-PT15M");
            icsBuilder.AppendLine("ACTION:DISPLAY");
            icsBuilder.AppendLine("DESCRIPTION:Reminder");
            icsBuilder.AppendLine("END:VALARM");
            icsBuilder.AppendLine("END:VEVENT");
            icsBuilder.AppendLine("END:VCALENDAR");

            return icsBuilder.ToString();
        }

        /// <summary>
        /// Try to parse the no separator time string to DateTime
        /// </summary>
        /// <param name="time">The specified DateTime string</param>
        /// <returns>Return the DateTime with instanceId specified format</returns>
        public static string ConvertInstanceIdFormat(string time)
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append(time.Substring(0, 4));
            stringBuilder.Append("-");
            stringBuilder.Append(time.Substring(4, 2));
            stringBuilder.Append("-");
            stringBuilder.Append(time.Substring(6, 5));
            stringBuilder.Append(":");
            stringBuilder.Append(time.Substring(11, 2));
            stringBuilder.Append(":");
            stringBuilder.Append(time.Substring(13, 2));
            stringBuilder.Append(".000");
            stringBuilder.Append(time.Substring(15));
            return stringBuilder.ToString();
        }
    }
}
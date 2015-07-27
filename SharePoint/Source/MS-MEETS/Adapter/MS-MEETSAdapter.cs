//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using System;
    using System.Net;
    using System.Web.Services.Protocols;

    /// <summary>
    /// MEETSAdapter implementation
    /// </summary>
    public partial class MS_MEETSAdapter : ManagedAdapterBase, IMS_MEETSAdapter
    {
        /// <summary>
        /// An instance of MeetingsSoap class, use to call the meetings web service.
        /// </summary>
        private MeetingsSoap service;

        /// <summary>
        /// Gets or sets the destination Url of web service operation.
        /// </summary>
        /// <value>Destination Url of web service operation.</value>
        public string Url
        {
            get
            {
                return this.service.Url;
            }

            set
            {
                this.service.Url = value;
            }
        }

        /// <summary>
        /// Overrides IAdapter's Initialize().
        /// </summary>
        /// <param name="testSite">A parameter represents an ITestSite instance.</param>
        public override void Initialize(ITestSite testSite)
        {
            base.Initialize(testSite);
            testSite.DefaultProtocolDocShortName = "MS-MEETS";

            // Initialize the MeetingSoap.
            this.service = Proxy.CreateProxy<MeetingsSoap>(this.Site);

            // Load common configuration.
            this.LoadCommonConfiguration();

            // Load SHOULDMAY configuration 
            this.LoadCurrentSutSHOULDMAYConfiguration();

            this.service.Url = Common.GetConfigurationPropertyValue("TargetServiceUrl", this.Site);
            string userName = Common.GetConfigurationPropertyValue("UserName", this.Site);
            string password = Common.GetConfigurationPropertyValue("Password", this.Site);
            string domain = Common.GetConfigurationPropertyValue("Domain", this.Site);
            this.service.Credentials = new NetworkCredential(userName, password, domain);
           
            this.SetSoapVersion(this.service);

            // When request Url include HTTPS prefix, avoid closing base connection.
            // Local client will accept all certificates after executing this function. 
            TransportProtocol transport = Common.GetConfigurationPropertyValue<TransportProtocol>("TransportType", this.Site);
            if (transport == TransportProtocol.HTTPS)
            {
                Common.AcceptServerCertificate();
            }

            // Configure the service timeout.
            string soapTimeOut = Common.GetConfigurationPropertyValue("ServiceTimeOut", this.Site);

            // 60000 means the configure SOAP Timeout is in milliseconds.
            this.service.Timeout = Convert.ToInt32(soapTimeOut) * 60000;
        }

        /// <summary>
        /// Adds a meeting to an existing workspace.
        /// </summary>
        /// <param name="organizerEmail">The e-mail address of the meeting organizer.</param>
        /// <param name="uid">A unique identifier represents this meeting</param>
        /// <param name="sequence">The sequence or revision number of this meeting instance. Null if not specified.</param>
        /// <param name="utcDateStamp">The time that this meeting instance was created.</param>
        /// <param name="title">The subject of this meeting.</param>
        /// <param name="location">The physical or virtual location in which this meeting is to take place.</param>
        /// <param name="utcDateStart">The time that this meeting begins.</param>
        /// <param name="utcDateEnd">The time that this meeting ends.</param>
        /// <param name="nonGregorian">Whether the meeting organizer is using the Gregorian calendar. Null if not specified.</param>
        /// <returns>The aggregation of AddMeetingResponseAddMeetingResult response or SoapException thrown.</returns>
        public SoapResult<AddMeetingResponseAddMeetingResult> AddMeeting(string organizerEmail, string uid, uint? sequence, DateTime? utcDateStamp, string title, string location, DateTime utcDateStart, DateTime utcDateEnd, bool? nonGregorian)
        {
            AddMeetingResponseAddMeetingResult result = null;
            SoapException exception = null;
            try
            {
                // Call AddMeeting method.
                result = this.service.AddMeeting(organizerEmail, uid, sequence ?? 0, sequence.HasValue, utcDateStamp ?? default(DateTime), utcDateStamp.HasValue, title, location, utcDateStart, utcDateEnd, nonGregorian ?? false, nonGregorian.HasValue);

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifyAddMeetingResponse(result);
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<AddMeetingResponseAddMeetingResult>(result, exception);
        }

        /// <summary>
        /// Adds a meeting to an existing workspace based on a calendar object.
        /// </summary>
        /// <param name="organizerEmail">The e-mail address of the meeting organizer.</param>
        /// <param name="icalText">Information about the meeting instance to add.</param>
        /// <returns>The aggregation of AddMeetingFromICalResponseAddMeetingFromICalResult response or SoapException thrown.</returns>
        public SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> AddMeetingFromICal(string organizerEmail, string icalText)
        {
            AddMeetingFromICalResponseAddMeetingFromICalResult result = null;
            SoapException exception = null;
            try
            {
                // Call AddMeetingFromIcal method.
                result = this.service.AddMeetingFromICal(organizerEmail, icalText);

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifyAddMeetingFromICalResponse(result);
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult>(result, exception);
        }

        /// <summary>
        /// Creates a new meeting workspace subsite.
        /// </summary>
        /// <param name="title">The title of the new meeting workspace.</param>
        /// <param name="templateName">The name of the template to use for the new meeting workspace.</param>
        /// <param name="lcid">The locale ID for the new workspace. Null if not specified.</param>
        /// <param name="timeZoneInformation">The time zone on the system of the meeting organizer.</param>
        /// <returns>The aggregation of CreateWorkspaceResponseCreateWorkspaceResult response or SoapException thrown.</returns>
        public SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> CreateWorkspace(string title, string templateName, uint? lcid, TimeZoneInf timeZoneInformation)
        {
            CreateWorkspaceResponseCreateWorkspaceResult result = null;
            SoapException exception = null;
            try
            {
                // Call CreateWorkspace method.
                result = this.service.CreateWorkspace(title, templateName, lcid ?? 0, lcid.HasValue, timeZoneInformation);

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifyCreateWorkspaceResponse(result);
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<CreateWorkspaceResponseCreateWorkspaceResult>(result, exception);
        }

        /// <summary>
        /// Deletes a workspace.
        /// </summary>
        /// <returns>The aggregation of empty response or SoapException thrown.</returns>
        public SoapResult<Null> DeleteWorkspace()
        {
            SoapException exception = null;
            try
            {
                // call DeleteWorkspace method.
                this.service.DeleteWorkspace();

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifyDeleteWorkspaceResponse();
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<Null>(Null.Value, exception);
        }

        /// <summary>
        /// Queries information from a Web site about meeting workspaces on it or information about a meeting workspace itself.
        /// </summary>
        /// <param name="requestFlags">Bit flags that specify what information to request from the Web site. Null if not specified.</param>
        /// <param name="lcid">The locale ID of the meeting workspace templates to return. Null if not specified.</param>
        /// <returns>The aggregation of GetMeetingsInformationResponseGetMeetingsInformationResult response or SoapException thrown.</returns>
        public SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> GetMeetingsInformation(MeetingInfoTypes? requestFlags, uint? lcid)
        {
            GetMeetingsInformationResponseGetMeetingsInformationResult result = null;
            SoapException exception = null;
            try
            {
                // Call GetMeetingsInformation method.
                result = this.service.GetMeetingsInformation((uint)(requestFlags ?? 0), requestFlags.HasValue, lcid ?? 0, lcid.HasValue);

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifyGetMeetingsInformationResponse(result, requestFlags);
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult>(result, exception);
        }

        /// <summary>
        /// Gets a list of created meeting workspace subsites from a parent web.
        /// </summary>
        /// <param name="recurring">Whether the meeting workspaces returned are limited to those that can be associated with a recurring meeting. Null if not specified.</param>
        /// <returns>The aggregation of GetMeetingWorkspacesResponseGetMeetingWorkspacesResult response or SoapException thrown.</returns>
        public SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult> GetMeetingWorkspaces(bool? recurring)
        {
            GetMeetingWorkspacesResponseGetMeetingWorkspacesResult result = null;
            SoapException exception = null;
            try
            {
                // Call GetMeetingWorkspaces method.
                result = this.service.GetMeetingWorkspaces(recurring ?? false, recurring.HasValue);

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifyGetMeetingWorkspacesResponse(result);
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult>(result, exception);
        }

        /// <summary>
        /// Deletes a meeting from an existing workspace.
        /// </summary>
        /// <param name="recurrenceId">The particular instance of a recurring meeting to delete. Null if not specified.</param>
        /// <param name="uid">A unique identifier represents the meeting to delete.</param>
        /// <param name="sequence">The sequence number of a meeting instance to delete. Null if not specified.</param>
        /// <param name="utcDateStamp">The time stamp for when this meeting instance was deleted. Null if not specified.</param>
        /// <param name="cancelMeeting">Whether the meeting is being cancelled in addition to being deleted from the workspace. Null if not specified.</param>
        /// <returns>The aggregation of empty response or SoapException thrown.</returns>
        public SoapResult<Null> RemoveMeeting(uint? recurrenceId, string uid, uint? sequence, DateTime? utcDateStamp, bool? cancelMeeting)
        {
            SoapException exception = null;
            try
            {
                // Call RemoveMeeting method.
                this.service.RemoveMeeting(recurrenceId ?? 0, recurrenceId.HasValue, uid, sequence ?? 0, sequence.HasValue, utcDateStamp ?? default(DateTime), utcDateStamp.HasValue, cancelMeeting ?? false, cancelMeeting.HasValue);

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifyRemoveMeetingResponse();
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<Null>(Null.Value, exception);
        }

        /// <summary>
        /// Restores a previously deleted meeting to a workspace.
        /// </summary>
        /// <param name="uid">A unique identifier represents the meeting to restore.</param>
        /// <returns>The aggregation of empty response or SoapException thrown.</returns>
        public SoapResult<Null> RestoreMeeting(string uid)
        {
            SoapException exception = null;
            try
            {
                // Call RestoreMeeting method.
                this.service.RestoreMeeting(uid);

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifyRestoreMeetingResponse();
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<Null>(Null.Value, exception);
        }

        /// <summary>
        /// Specifies attendee response to a meeting request in a workspace.
        /// </summary>
        /// <param name="attendeeEmail">The e-mail address of the attendee responding.</param>
        /// <param name="recurrenceId">The particular instance of a recurring meeting associated with the response. Null if not specified.</param>
        /// <param name="uid">A unique identifier represents the meeting associated with the response.</param>
        /// <param name="sequence">The sequence number of a meeting instance associated with the response. Null if not specified.</param>
        /// <param name="utcDateTimeOrganizerCriticalChange">The time stamp at which the attendee responded to the meeting, translated to the time zone of the meeting organizer. Null if not specified.</param>
        /// <param name="utcDateTimeAttendeeCriticalChange">The time stamp that contains the time at which the attendee responded to the meeting, translated to the time zone of the attendee. Null if not specified.</param>
        /// <param name="response">Attendee acceptance or rejection of the meeting invitation. Null if not specified.</param>
        /// <returns>The aggregation of empty response or SoapException thrown.</returns>
        public SoapResult<Null> SetAttendeeResponse(string attendeeEmail, uint? recurrenceId, string uid, uint? sequence, DateTime? utcDateTimeOrganizerCriticalChange, DateTime? utcDateTimeAttendeeCriticalChange, AttendeeResponse? response)
        {
            SoapException exception = null;
            try
            {
                // Call SetAttendeeResponse method.
                this.service.SetAttendeeResponse(
                    attendeeEmail,
                    recurrenceId ?? 0,
                    recurrenceId.HasValue,
                    uid,
                    sequence ?? 0,
                    sequence.HasValue,
                    utcDateTimeOrganizerCriticalChange ?? default(DateTime),
                    utcDateTimeOrganizerCriticalChange.HasValue,
                    utcDateTimeAttendeeCriticalChange ?? default(DateTime),
                    utcDateTimeAttendeeCriticalChange.HasValue,
                    response ?? 0,
                    response.HasValue);

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifySetAttendeeResponseResponse();
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<Null>(Null.Value, exception);
        }

        /// <summary>
        /// Specifies a user friendly name for a workspace.
        /// </summary>
        /// <param name="title">The new title of the meeting workspace.</param>
        /// <returns>The aggregation of empty response or SoapException thrown.</returns>
        public SoapResult<Null> SetWorkspaceTitle(string title)
        {
            SoapException exception = null;
            try
            {
                // Call SetWorkspaceTitle method.
                this.service.SetWorkspaceTitle(title);

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifySetWorkspaceTitleResponse();
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<Null>(Null.Value, exception);
        }

        /// <summary>
        /// Updates content of a meeting in a workspace.
        /// </summary>
        /// <param name="uid">A unique identifier represents the meeting to update.</param>
        /// <param name="sequence">The updated sequence number of this meeting. Null if not specified.</param>
        /// <param name="utcDateStamp">The time at which the meeting is being updated. Null if not specified.</param>
        /// <param name="title">The updated subject of this meeting.</param>
        /// <param name="location">The updated physical or virtual location in which this meeting is to take place.</param>
        /// <param name="utcDateStart">The updated beginning time of this meeting.</param>
        /// <param name="utcDateEnd">The updated end time of this meeting.</param>
        /// <param name="nonGregorian">Whether the updated meeting is not in the Gregorian calendar. Null if not specified.</param>
        /// <returns>The aggregation of empty response or SoapException thrown.</returns>
        public SoapResult<Null> UpdateMeeting(string uid, uint? sequence, DateTime? utcDateStamp, string title, string location, DateTime utcDateStart, DateTime utcDateEnd, bool? nonGregorian)
        {
            SoapException exception = null;
            try
            {
                // Call UpdateMeeting method.
                this.service.UpdateMeeting(
                    uid,
                    sequence ?? 0,
                    sequence.HasValue,
                    utcDateStamp ?? default(DateTime),
                    utcDateStamp.HasValue,
                    title,
                    location,
                    utcDateStart,
                    utcDateEnd,
                    nonGregorian ?? false,
                    nonGregorian.HasValue);

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifyUpdateMeetingResponse();
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<Null>(Null.Value, exception);
        }

        /// <summary>
        /// Updates content of a meeting in a workspace base on a calendar object.
        /// </summary>
        /// <param name="icalText">Updated information about the meeting instance.</param>
        /// <param name="ignoreAttendees">Whether this is a scheduling-only update, or an update that affects attendees. Null if not specified.</param>
        /// <returns>The aggregation of UpdateMeetingFromICalResponseUpdateMeetingFromICalResultUpdateMeetingFromICal response or SoapException thrown.</returns>
        public SoapResult<UpdateMeetingFromICalResponseUpdateMeetingFromICalResult> UpdateMeetingFromICal(string icalText, bool? ignoreAttendees)
        {
            UpdateMeetingFromICalResponseUpdateMeetingFromICalResult result = null;
            SoapException exception = null;
            try
            {
                // Call UpdateMeetingFromICal method.
                result = this.service.UpdateMeetingFromICal(icalText, ignoreAttendees ?? false, ignoreAttendees.HasValue);

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();
                this.ValidateAndCaptureCommonMessageSyntax();
                this.VerifyUpdateMeetingFromICalResponse(result);
            }
            catch (SoapException ex)
            {
                exception = ex;

                // As response successfully returned, the transport related requirements can be captured.
                this.CaptureTransportRelatedRequirements();

                // Validate soap fault message structure and capture related requirements.
                this.ValidateAndCaptureSoapFaultRequirements(exception);
            }

            return new SoapResult<UpdateMeetingFromICalResponseUpdateMeetingFromICalResult>(result, exception);
        }

        /// <summary>
        /// A method used to load Common Configuration
        /// </summary>
        private void LoadCommonConfiguration()
        {
            // Merge the common configuration into local configuration
            string conmmonConfigFileName = Common.GetConfigurationPropertyValue("CommonConfigurationFileName", this.Site);

            // Execute the merge the common configuration
            Common.MergeGlobalConfig(conmmonConfigFileName, this.Site);

            Common.CheckCommonProperties(this.Site, true);
        }

        /// <summary>
        /// A method used to load SHOULDMAY Configuration according to the current SUT version
        /// </summary>
        private void LoadCurrentSutSHOULDMAYConfiguration()
        {
            Common.MergeSHOULDMAYConfig(this.Site);
        }

        /// <summary>
        /// Set the SOAP version according to the SoapVersion property.
        /// </summary>
        /// <param name="meetingProxy">set meeting proxy</param>
        private void SetSoapVersion(MeetingsSoap meetingProxy)
        {
            SoapVersion soapVersion = Common.GetConfigurationPropertyValue<SoapVersion>("SoapVersion", this.Site);

            switch (soapVersion)
            {
                case SoapVersion.SOAP11:
                    {
                        meetingProxy.SoapVersion = SoapProtocolVersion.Soap11;
                        break;
                    }

                default:
                    {
                        meetingProxy.SoapVersion = SoapProtocolVersion.Soap12;
                        break;
                    }
            }
        }
    }
}
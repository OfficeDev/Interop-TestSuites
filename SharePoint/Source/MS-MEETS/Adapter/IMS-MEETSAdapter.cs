namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using System;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// This interface defines an Adapter which contains web service operations used to communicate with SUT.
    /// </summary>
    public interface IMS_MEETSAdapter : IAdapter
    {
        /// <summary>
        /// Gets or sets the entry point Url of web service operation.
        /// </summary>
        /// <value>The service Url</value>
        string Url
        {
            get;
            set;
        }

        /// <summary>
        /// Adds a meeting to an existing workspace.
        /// </summary>
        /// <param name="organizerEmail">The e-mail address of the meeting organizer.</param>
        /// <param name="uid">A unique identifier represents this meeting</param>
        /// <param name="sequence">The sequence or revision number of this meeting instance. Null if not specified.</param>
        /// <param name="utcDateStamp">Represents the time when this meeting instance was created.</param>
        /// <param name="title">The subject of this meeting.</param>
        /// <param name="location">The physical or virtual location in which this meeting is to take place.</param>
        /// <param name="utcDateStart">Represents the time when this meeting instance was begins.</param>
        /// <param name="utcDateEnd">Represents the time at which this meeting instance was ends.</param>
        /// <param name="nonGregorian">Whether the meeting organizer is using the Gregorian calendar. Null if not specified.</param>
        /// <returns>The aggregation of AddMeetingResponseAddMeetingResult response or SoapException thrown.</returns>
        SoapResult<AddMeetingResponseAddMeetingResult> AddMeeting(string organizerEmail, string uid, uint? sequence, DateTime? utcDateStamp, string title, string location, DateTime utcDateStart, DateTime utcDateEnd, bool? nonGregorian);

        /// <summary>
        /// Adds a meeting to an existing workspace based on a calendar object.
        /// </summary>
        /// <param name="organizerEmail">The e-mail address of the meeting organizer.</param>
        /// <param name="icalText">Information about the meeting instance to add.</param>
        /// <returns>The aggregation of AddMeetingFromICalResponseAddMeetingFromICalResult response or SoapException thrown.</returns>
        SoapResult<AddMeetingFromICalResponseAddMeetingFromICalResult> AddMeetingFromICal(string organizerEmail, string icalText);

        /// <summary>
        /// Creates a new meeting workspace subsite.
        /// </summary>
        /// <param name="title">The title of the new meeting workspace.</param>
        /// <param name="templateName">The name of the template to use for the new meeting workspace.</param>
        /// <param name="lcid">The locale ID for the new workspace. Null if not specified.</param>
        /// <param name="timeZoneInformation">The time zone on the system of the meeting organizer.</param>
        /// <returns>The aggregation of CreateWorkspaceResponseCreateWorkspaceResult response or SoapException thrown.</returns>
        SoapResult<CreateWorkspaceResponseCreateWorkspaceResult> CreateWorkspace(string title, string templateName, uint? lcid, TimeZoneInf timeZoneInformation);

        /// <summary>
        /// Deletes a workspace.
        /// </summary>
        /// <returns>The aggregation of empty response or SoapException thrown</returns>
        SoapResult<Null> DeleteWorkspace();

        /// <summary>
        /// Queries information from a Web site about meeting workspaces on it or information about a meeting workspace itself.
        /// </summary>
        /// <param name="requestFlags">Bit flags that specify what information to request from the Web site. Null if not specified.</param>
        /// <param name="lcid">The locale ID of the meeting workspace templates to return. Null if not specified.</param>
        /// <returns>The aggregation of GetMeetingsInformationResponseGetMeetingsInformationResult response or SoapException thrown.</returns>
        SoapResult<GetMeetingsInformationResponseGetMeetingsInformationResult> GetMeetingsInformation(MeetingInfoTypes? requestFlags, uint? lcid);

        /// <summary>
        /// Gets a list of created meeting workspace subsites from a parent web site.
        /// </summary>
        /// <param name="recurring">Whether the meeting workspaces returned are limited to those that can be associated with a recurring meeting. Null if not specified.</param>
        /// <returns>The aggregation of GetMeetingWorkspacesResponseGetMeetingWorkspacesResult response or SoapException thrown.</returns>
        SoapResult<GetMeetingWorkspacesResponseGetMeetingWorkspacesResult> GetMeetingWorkspaces(bool? recurring);

        /// <summary>
        /// Deletes a meeting from an existing workspace.
        /// </summary>
        /// <param name="recurrenceId">The particular instance of a recurring meeting to delete. Null if not specified.</param>
        /// <param name="uid">A unique identifier represents the meeting to delete.</param>
        /// <param name="sequence">The sequence number of a meeting instance to delete. Null if not specified.</param>
        /// <param name="utcDateStamp">The time stamp for when this meeting instance was deleted. Null if not specified.</param>
        /// <param name="cancelMeeting">Whether the meeting is being cancelled in addition to being deleted from the workspace. Null if not specified.</param>
        /// <returns>The aggregation of empty response or SoapException thrown.</returns>
        SoapResult<Null> RemoveMeeting(uint? recurrenceId, string uid, uint? sequence, DateTime? utcDateStamp, bool? cancelMeeting);

        /// <summary>
        /// Restores a previously deleted meeting to a workspace.
        /// </summary>
        /// <param name="uid">A unique identifier represents the meeting to restore.</param>
        /// <returns>The aggregation of empty response or SoapException thrown.</returns>
        SoapResult<Null> RestoreMeeting(string uid);

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
        SoapResult<Null> SetAttendeeResponse(string attendeeEmail, uint? recurrenceId, string uid, uint? sequence, DateTime? utcDateTimeOrganizerCriticalChange, DateTime? utcDateTimeAttendeeCriticalChange, AttendeeResponse? response);

        /// <summary>
        /// Specifies a user friendly name for a workspace.
        /// </summary>
        /// <param name="title">The new title of the meeting workspace.</param>
        /// <returns>The aggregation of empty response or SoapException thrown.</returns>
        SoapResult<Null> SetWorkspaceTitle(string title);

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
        SoapResult<Null> UpdateMeeting(string uid, uint? sequence, DateTime? utcDateStamp, string title, string location, DateTime utcDateStart, DateTime utcDateEnd, bool? nonGregorian);

        /// <summary>
        /// Updates content of a meeting in a workspace base on a calendar object.
        /// </summary>
        /// <param name="icalText">Updated information about the meeting instance.</param>
        /// <param name="ignoreAttendees">Whether this is a scheduling-only update, or an update that affects attendees. Null if not specified.</param>
        /// <returns>The aggregation of UpdateMeetingFromICalResponseUpdateMeetingFromICalResultUpdateMeetingFromICal response or SoapException thrown.</returns>
        SoapResult<UpdateMeetingFromICalResponseUpdateMeetingFromICalResult> UpdateMeetingFromICal(string icalText, bool? ignoreAttendees);
    }
}
namespace Microsoft.Protocols.TestSuites.MS_MEETS
{
    using System;
    using System.Text;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;

    /// <summary>
    /// A class contains all helper methods used in test case level.
    /// </summary>
    public static class TestSuiteBase
    {
        /// <summary>
        /// A random generator using current time seeds.
        /// </summary>
        private static Random random;

        /// <summary>
        /// A ITestSite instance
        /// </summary>
        private static ITestSite testSite;

        /// <summary>
        /// A method used to initialize the TestSuiteBase with specified ITestSite instance.
        /// </summary>
        /// <param name="site">A parameter represents ITestSite instance.</param>
        public static void Initialize(ITestSite site)
        {
            testSite = site;
        }

        /// <summary>
        /// Generate an iCalendar with specified UID and attendees.
        /// </summary>
        /// <param name="uid">The unique id</param>
        /// <param name="isRecurring">Specifies whether the calendar is a recurring calendar</param>
        /// <param name="attendees">The collection of attendees</param>
        /// <returns>A string contains the iCalendar</returns>
        public static string GetICalendar(string uid, bool isRecurring, params string[] attendees)
        {
            // iCalendar Tag  DTSTAMP specifies the date/time that the instance of the iCalendar object was created.
            // iCalendar Tag  DESCRIPTION specifies the description of the meeting.
            // iCalendar Tag  LOCATION defines the intended venue for the meeting.
            // iCalendar Tag  SUMMARY defines a short summary or subject for the meeting.
            // iCalendar Tag  DTSTART specifies when the meeting begins. 
            string meetingTitle = TestSuiteBase.GetUniqueMeetingTitle();
            string meetingLocation = TestSuiteBase.GetUniqueMeetingLocation();
            string icalTemplate = "BEGIN:VCALENDAR\n"
                                + "BEGIN:VEVENT\n"
                                + "{0}"
                                + "DTSTAMP:20090314T143000Z\n"
                                + "DESCRIPTION:" + meetingTitle + "\n"
                                + "SUMMARY:" + meetingTitle + "\n"
                                + "DTSTART;TZID=US-Eastern:19970902T090000\n"
                                + "LOCATION:" + meetingLocation + "\n"
                                + "END:VEVENT\n"
                                + "END:VCALENDAR\n";

            StringBuilder stringBuilder = new StringBuilder();

            // iCalendar Tag  UID specifies an unique UID for the meeting.
            stringBuilder.AppendLine(string.Format("UID:{0}", uid));
            foreach (string attendee in attendees)
            {
                // iCalendar Tag  ATTENDEE specifies attendees of the meeting. 
                stringBuilder.AppendLine(string.Format("ATTENDEE:MAILTO:{0}", attendee));
            }

            if (isRecurring)
            {
                // iCalendar Tag  RRULE specifies the recurring rule of this meeting.
                stringBuilder.AppendLine("RRULE:FREQ=DAILY;COUNT=5");
            }

            string icalendar = string.Format(icalTemplate, stringBuilder.ToString());
            return icalendar;
        }

        /// <summary>
        /// Generate an iCalendar with specified UID and attendees.
        /// </summary>
        /// <param name="uid">The unique id</param>
        /// <param name="isRecurring">Specifies whether the calendar is a recurring calendar</param>
        /// <param name="meetingTitle">The title of the meeting.</param>
        /// <param name="meetingLocation">The location of the meeting.</param>
        /// <param name="attendees">The collection of attendees</param>
        /// <returns>A string contains the iCalendar</returns>
        public static string GetICalendar(
                                string uid, 
                                bool isRecurring, 
                                string meetingTitle, 
                                string meetingLocation, 
                                params string[] attendees)
        {
            // iCalendar Tag  DTSTAMP specifies the date/time that the instance of the iCalendar object was created.
            // iCalendar Tag  DESCRIPTION specifies the description of the meeting.
            // iCalendar Tag  LOCATION defines the intended venue for the meeting.
            // iCalendar Tag  SUMMARY defines a short summary or subject for the meeting.
            // iCalendar Tag  DTSTART specifies when the meeting begins. 
            string icalTemplate = "BEGIN:VCALENDAR\n"
                                + "BEGIN:VEVENT\n"
                                + "{0}"
                                + "DTSTAMP:20090314T143000Z\n"
                                + "DESCRIPTION:" + meetingTitle + "\n"
                                + "SUMMARY:" + meetingTitle + "\n"
                                + "DTSTART;TZID=US-Eastern:19970902T090000\n"
                                + "LOCATION:" + meetingLocation + "\n"
                                + "END:VEVENT\n"
                                + "END:VCALENDAR\n";

            StringBuilder stringBuilder = new StringBuilder();

            // iCalendar Tag  UID specifies an unique UID for the meeting.
            stringBuilder.AppendLine(string.Format("UID:{0}", uid));
            foreach (string attendee in attendees)
            {
                // iCalendar Tag  ATTENDEE specifies attendees of the meeting. 
                stringBuilder.AppendLine(string.Format("ATTENDEE:MAILTO:{0}", attendee));
            }

            if (isRecurring)
            {
                // iCalendar Tag  RRULE specifies the recurring rule of this meeting.
                stringBuilder.AppendLine("RRULE:FREQ=DAILY;COUNT=5");
            }

            string icalendar = string.Format(icalTemplate, stringBuilder.ToString());
            return icalendar;
        }

        /// <summary>
        /// A method is used to get a Unique Workspace title.
        /// </summary>
        /// <returns>A return value represents the unique Workspace title that is combined with the Workspace Object name and time stamp</returns>
        public static string GetUniqueWorkspaceTitle()
        {
            return Common.GenerateResourceName(testSite, "WorkspaceTitle");
        }

        /// <summary>
        /// A method is used to get a Unique Meeting title.
        /// </summary>
        /// <returns>A return value represents the unique Meeting title that is combined with the Meeting Object name and time stamp</returns>
        public static string GetUniqueMeetingTitle()
        {
            return Common.GenerateResourceName(testSite, "MeetingTitle");
        }

        /// <summary>
        /// A method is used to get a Unique Meeting location.
        /// </summary>
        /// <returns>A return value represents the unique Meeting location that is combined with the Location Object name and time stamp</returns>
        public static string GetUniqueMeetingLocation()
        {
            return Common.GenerateResourceName(testSite, "MeetingLocation");
        }

        /// <summary>
        /// This method is used to generate random string in the range A-Z with the specified string size.
        /// </summary>
        /// <param name="size">A parameter represents the generated string size.</param>
        /// <returns>Returns the random generated string.</returns>
        public static string GenerateRandomString(int size)
        {
            random = new Random((int)DateTime.Now.Ticks);
            StringBuilder builder = new StringBuilder();
            char ch;
            for (int i = 0; i < size; i++)
            {
                int intIndex = Convert.ToInt32(Math.Floor((26 * random.NextDouble()) + 65));
                ch = Convert.ToChar(intIndex);
                builder.Append(ch);
            }

            return builder.ToString();
        }

        /// <summary>
        /// This method is used to generate random attendees with the specified size.
        /// </summary>
        /// <param name="size">A parameter represents the generated attendees size.</param>
        /// <returns>Returns the random generated attendees.</returns>
        public static string[] GenerateAttendees(int size)
        {
            string[] attendees = new string[size];
            for (int i = 0; i < size; i++)
            {
                attendees[i] = GenerateRandomString(5) + "@contoso.com";
            }

            return attendees;
        }
    }
}
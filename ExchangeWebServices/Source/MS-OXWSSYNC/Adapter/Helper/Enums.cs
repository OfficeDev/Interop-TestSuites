namespace Microsoft.Protocols.TestSuites.MS_OXWSSYNC
{
    /// <summary>
    /// An enumeration that describes type of item. 
    /// </summary>
    public enum Item
    {
        /// <summary>
        /// Represent a meeting request
        /// </summary>
        MeetingRequest,

        /// <summary>
        /// Represent a meeting response
        /// </summary>
        MeetingResponse,

        /// <summary>
        /// Represent a meeting cancellation
        /// </summary>
        MeetingCancellation,

        /// <summary>
        /// Represent a calendar item
        /// </summary>
        CalendarItem
    }
}
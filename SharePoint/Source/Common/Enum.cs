namespace Microsoft.Protocols.TestSuites.Common
{
    /// <summary>
    /// The protocol transport type which is used to transfer messages between the client and SUT.
    /// </summary>
    public enum TransportProtocol
    {
        /// <summary>
        /// The transport is SOAP over HTTP.
        /// </summary>
        HTTP,

        /// <summary>
        /// The transport is SOAP over HTTPS.
        /// </summary>
        HTTPS
    }

    /// <summary>
    /// The SOAP version which is used to format the messages between the client and SUT.
    /// </summary>
    public enum SoapVersion
    {
        /// <summary>
        /// The messages are formatted with SOAP 1.1.
        /// </summary>
        SOAP11,

        /// <summary>
        /// The messages are formatted with SOAP 1.2.
        /// </summary>
        SOAP12
    }

    /// <summary>
    /// The version of SUT.
    /// </summary>
    public enum SutVersion
    {
        /// <summary>
        /// The SUT is Windows SharePoint Services 3.0 SP3.
        /// </summary>
        WindowsSharePointServices3,

        /// <summary>
        /// The SUT is Microsoft SharePoint Foundation 2010 SP2.
        /// </summary>
        SharePointFoundation2010,

        /// <summary>
        /// The SUT is Microsoft SharePoint Foundation 2013 SP1.
        /// </summary>
        SharePointFoundation2013,

        /// <summary>
        /// The SUT is Microsoft Office SharePoint Server 2007 SP3.
        /// </summary>
        SharePointServer2007,

        /// <summary>
        /// The SUT is Microsoft SharePoint Server 2010 SP2.
        /// </summary>
        SharePointServer2010,

        /// <summary>
        /// The SUT is Microsoft SharePoint Server 2013 SP1.
        /// </summary>
        SharePointServer2013,

        /// <summary>
        /// The SUT is Microsoft SharePoint Server 2016.
        /// </summary>
        SharePointServer2016,

        /// <summary>
        /// The SUT is Microsoft SharePoint Server 2019.
        /// </summary>
        SharePointServer2019,

        /// <summary>
        /// The SUT is Microsoft SharePoint Server Subscription Edition Preview.
        /// </summary>
        SharePointServerSubscriptionEditionPreview
    }

    /// <summary>
    /// Represent Result of Validation
    /// </summary>
    public enum ValidationResult
    {
        /// <summary>
        /// Indicate the validation is success.
        /// </summary>
        Success,

        /// <summary>
        /// Indicate the validation is error.
        /// </summary>
        Error,

        /// <summary>
        /// Indicate the validation is warning.
        /// </summary>
        Warning,

        /// <summary>
        /// Indicate the validation is inconclusive.
        /// </summary>
        Inconclusive
    }

    /// <summary>
    /// The server's list template types
    /// </summary>
    public enum TemplateType
    {
        /// <summary>
        /// Unknown Template
        /// </summary>
        Unkown = 0,

        /// <summary>
        /// Invalid Template
        /// </summary>
        Invalid = -1,

        /// <summary>
        /// Generic List Template.
        /// </summary>
        Generic_List = 100,

        /// <summary>
        /// Document Library Template.
        /// </summary>
        Document_Library = 101,

        /// <summary>
        /// Discussion Template
        /// </summary>
        Discussion_Board = 108,

        /// <summary>
        /// Issues Template
        /// </summary>
        Issues = 1100,

        /// <summary>
        /// Survey Template
        /// </summary>
        Survey = 102,

        /// <summary>
        /// Links Template
        /// </summary>
        Links = 103,

        /// <summary>
        /// Announcements Template
        /// </summary>
        Announcements = 104,

        /// <summary>
        /// Contacts Template
        /// </summary>
        Contacts = 105,

        /// <summary>
        /// Events Template
        /// </summary>
        Events = 106,

        /// <summary>
        /// Tasks Template
        /// </summary>
        Tasks = 107,

        /// <summary>
        /// Image Library Template
        /// </summary>
        Image = 109,

        /// <summary>
        /// Data Sources Template
        /// </summary>
        DataSource = 110,

        /// <summary>
        /// User Info Catalog Template
        /// </summary>
        UserInfo = 112,

        /// <summary>
        /// Web Part Catalog Template
        /// </summary>
        WebPartCatalog = 113,

        /// <summary>
        /// XML Form Template
        /// </summary>
        XMLForm = 115,

        /// <summary>
        /// Master Page Catalog Template
        /// </summary>
        MasterPageCatalog = 116,

        /// <summary>
        /// No Code Workflows Template
        /// </summary>
        NoCodeWorkflow = 117,

        /// <summary>
        /// Workflow Process Template
        /// </summary>
        WorkflowProcess = 118,

        /// <summary>
        /// Webpage Library Template
        /// </summary>
        WebPage = 119,

        /// <summary>
        /// Custom Grid Template
        /// </summary>
        Grid = 120,

        /// <summary>
        /// Data Connection Library Template
        /// </summary>
        DataConnection = 130,

        /// <summary>
        /// Work Flow History Template
        /// </summary>
        WorkflowHistory = 140,

        /// <summary>
        /// Gantt Tasks Template
        /// </summary>
        Gantt_Task = 150,

        /// <summary>
        /// Meetings Template
        /// </summary>
        Meetings = 200,

        /// <summary>
        /// Agenda Template
        /// </summary>
        Agenda = 201,

        /// <summary>
        /// Meeting User Template
        /// </summary>
        MeetingUser = 202,

        /// <summary>
        /// Decision (Meeting) Template
        /// </summary>
        Decision = 204,

        /// <summary>
        /// Meeting Objective Template
        /// </summary>
        MeetingObjective = 207,

        /// <summary>
        /// Textbox Template
        /// </summary>
        Textbox = 210,

        /// <summary>
        /// Things To Bring (Meeting) Template
        /// </summary>
        ThingsToBring = 211,

        /// <summary>
        /// Homepage Library Template
        /// </summary>
        HomePageLibrary = 212,

        /// <summary>
        /// Posts (Blog) Template
        /// </summary>
        Posts = 301,

        /// <summary>
        /// Comments (Blog) Template
        /// </summary>
        Comments = 302,

        /// <summary>
        /// Categories (Blog) Template
        /// </summary>
        Categories = 303,

        /// <summary>
        /// Resources List Template
        /// </summary>
        ResourcesList = 402,
    }
}
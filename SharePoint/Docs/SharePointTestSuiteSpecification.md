SharePoint Test Suites Specification
=====================================================================================================================================================================================================================================================================================================================================================================================================================================================================

- [Introduction](#introduction)

- [Requirement specification](#requirement-specification)

- [Design considerations](#design-considerations)

- [Package design](#package-design)

Introduction
=====================================================================================================================================================================================================================================================================================================================================================================================================================================================================

The SharePoint Test Suites are implemented as synthetic clients running
against a server-side implementation of a given SharePoint protocol.
They are designed in a client-to-server relationship and were originally
developed for the in-house testing of the Microsoft Open Specifications.

Microsoft Open Specifications were written using the normative language
defined in
[RFC2119](http://go.microsoft.com/fwlink/?LinkId=117453);from which
statements are extracted as protocol requirements to be listed in the
requirement specification.See [Requirement
Specification](#requirement-specification). This document describes how
the SharePoint Protocol Test Suites are designed to verify that the
server behaves in compliance with normative protocol requirements in the
technical specification.

In a single test suite, similar or related requirements are grouped into
one test case. Test cases on the same command or operation are grouped
into one scenario.

The technical specifications listed in the following table are included
in the SharePoint Protocol Test Suites package. The version of these
technical specifications is v20160715.

SharePoint Protocol technical specifications

  Technical specification  | Protocol name
 :------------- | :------------- 
  MS-ADMINS        |  		  [Administration Web Service Protocol](https://go.microsoft.com/fwlink/?linkid=835091)
  MS-AUTHWS        |       [Authentication Web Service Protocol](https://go.microsoft.com/fwlink/?linkid=835092)
  MS-COPYS         |         [Copy Web Service Protocol](https://go.microsoft.com/fwlink/?linkid=835093)
  MS-CPSWS         |         [SharePoint Claim Provider Web Service Protocol](https://go.microsoft.com/fwlink/?linkid=835094)
  MS-DWSS          |         [Document Workspace Web Service Protocol](http://go.microsoft.com/fwlink/?LinkId=255888)
  MS-LISTSWS       |         [Lists Web Service Protocol](http://go.microsoft.com/fwlink/?LinkId=255885)
  MS-MEETS         |         [Meetings Web Services Protocol](https://go.microsoft.com/fwlink/?linkid=835095)
  MS-OFFICIALFILE  |         [Official File Web Service Protocol](https://go.microsoft.com/fwlink/?linkid=835096)
  MS-OUTSPS        |         [Lists Client Sync Protocol](https://go.microsoft.com/fwlink/?linkid=835097)
  MS-SHDACCWS      |         [Shared Access Web Service Protocol](https://go.microsoft.com/fwlink/?linkid=835098)
  MS-SITESS        |         [Sites Web Service Protocol](http://go.microsoft.com/fwlink/?LinkId=255887)
  MS-VERSS         |         [Versions Web Service Protocol](http://go.microsoft.com/fwlink/?LinkId=255886)
  MS-VIEWSS        |         [Views Web Service Protocol](https://go.microsoft.com/fwlink/?linkid=835099)
  MS-WDVMODUU      |         [Office Document Update Utility Extensions](https://go.microsoft.com/fwlink/?linkid=835100)
  MS-WEBSS         |         [Webs Web Service Protocol](https://go.microsoft.com/fwlink/?linkid=835101)
  MS-WSSREST       |         [ListData Data Service Protocol](https://go.microsoft.com/fwlink/?linkid=835102)
  MS-WWSP          |         [Workflow Web Service Protocol](https://go.microsoft.com/fwlink/?linkid=835103)

Requirement specification <a name="requirement-specification"></a>
======================================================================================================================================================================================================================================================================

A requirement specification contains a list of requirements that are
extracted from statements in the technical specification. Each technical
specification has one corresponding requirement specification named as
MS-XXXX\_RequirementSpecification.xlsx, which can be found in the
Docs\\MS-XXXX folder in the SharePoint Protocol Test Suites package
together with the technical specification.

The requirements are categorized as normative or informative. If the
statement of the requirement is required for interoperability, the
requirement is normative. If the statement of the requirement is
clarifying information or high-level introduction, and removal of it
does not affect interoperability, the requirement is informative.

Each requirement applies to a specific scope: server, client, or both.
If the requirement describes a behavior performed by the responder, the
scope of the requirement is server. If the requirement describes a
behavior performed by the initiator, the scope of the requirement is
client. If the requirement describes a behavior performed by both
initiator and responder, the scope of the requirement is both.

The test suites cover normative requirements which describes a behavior
performed by the responder. For a detailed requirements list and
classification, see the MS-XXXX\_RequirementSpecification.xlsx.

Design considerations
=====================

Assumptions
-----------

-   The SharePoint Protocol Test Suites are not designed to run
    multi-protocol user scenarios, but rather provide a way to exercise
    certain operations documented in a technical specification.

-   The test suites are functional tests that verify the compatibility
    of the system under test (SUT) with a protocol implementation.

-   The test suites do not cover every protocol requirement and in no
    way certify an implementation, even if all tests pass.

-   The test suites verify the server-side testable requirements; they
    do not verify the requirements related to client behaviors and
    server internal behaviors.

Dependencies
------------

-   All SharePoint Protocol Test Suites depend on the Protocol Test
    Framework (PTF) to derive managed adapters.

-   All SharePoint Protocol Test Suites depends on the SOAP messaging
    protocol for exchanging structured data and type information.

-   All SharePoint Protocol Test Suites depends on HTTP protocol or
    HTTPS protocol to transmit the messages.

-   All SharePoint Protocol Test Suite depends on the wsdl.exe tool in
    the .NET Framework SDK to generate the proxy class.

Package design
==============

SharePoint Protocol Test Suites are implemented as synthetic clients
running against a server-side implementation of a given SharePoint
protocol. The test suites verify the server-side and testable
requirements.

Architecture
------------

The following figure illustrates the SharePoint Protocol Test Suites
architecture.

![alt tag] (./Doc-Images/SharePoint_Spec_Architecture.png)

**Figure 1: Architecture**

The following outlines the details of the test suites architecture:

**SUT**

The system under test (SUT) hosts the server-side implementation of the
protocol, which test suites run against.

-   From a third-party’s point of view, the SUT is a
    server implementation.

-   The following products have been tested with the test suites on the
    Windows platform.

    -   Windows SharePoint Services 3.0 Service Pack 3 (SP3)

    -   Microsoft SharePoint Foundation 2010 Service Pack 2 (SP2)

    -   Microsoft SharePoint Foundation 2013 SP1

    -   Microsoft Office SharePoint Server 2007 Service Pack 3 (SP3)

    -   Microsoft SharePoint Server 2010 Service Pack 2 (SP2)

    -   Microsoft SharePoint Server 2013 SP1

    -   Microsoft SharePoint Server 2016

**Test Suite Client**

The test suites act as synthetic clients to communicate with the SUT and
validate the requirements gathered from technical specifications. The
SharePoint Protocol Test Suites include one common library, 17 adapters
and 17 test suites.

-   The test suites communicate with SUT via a protocol adapter and SUT
    control adapter to verify if the SUT behaves in the way that is
    compliant with normative protocol requirements.

Common library
--------------

The common library provides implementation of helper methods.

### Helper methods

The common library defines a series of helper methods. The helper
methods can be classified into following categories.

-   Access the properties in the configuration file.

-   Generate resource name.

-   Schema validation.

-   Other methods which are used by multiple test suites.

Adapter
-------

Adapters are interfaces between the test suites and the SUT. There are
two types of adapter: protocol adapter and SUT control adapter. In most
cases, modifications to the protocol adapter will not be required for
non-Microsoft SUT implementations. However, the SUT control adapter
should be appropriately configured to connect to a non-Microsoft SUT
implementation. All test suites in the package contain a protocol
adapter, six of them contain a SUT control adapter.

### Protocol Adapter

The protocol adapter is a managed adapter, which is derived from the
ManagedAdapterBase class in the PTF. It provides an interface that is
used by the test cases to construct protocol request messages that will
be sent to the SUT. The protocol adapter also acts as an intermediary
between the test cases and the transport classes, receiving messages,
sending messages, parsing responses from the transport classes, and
validating the SUT response according to the normative requirement in
the technical specification.

All protocol adapters use the proxy class of each protocol to send and
receive messages.

### SUT Control Adapter 

The SUT control adapter manages all the control functions of the test
suites that are not associated with the protocol. For example, the setup
and teardown are managed through the SUT control adapter. The SUT
control adapter is designed to work with the Microsoft implementation of
the SUT. However, it is configurable to allow the test suites to run
against non-Microsoft implementations of the SUT.

There are 15 protocols that have a SUT control adapter in the SharePoint
Protocol test suites package: MS-ADMINS, MS-COPYS, MS-CPSWS, MS-DWSS,
MS-LISTSWS, MS-MEETS, MS-OFFICIALFILE, MS-OUTSPS, MS-SHDACCWS,
MS-SITESS, MS-VERSS, MS-VIEWSS, MS-WEBSS, MS-WSSREST and MS-WWSP.

Test suite
----------

The test suites verify the server-side and testable requirements listed
in the requirement specification. The test suites call the protocol
adapter to send and receive message between the protocol adapter and the
SUT, and call the SUT control adapter to change the SUT state. The test
suites consists of a series test cases which are categorized to several
scenarios. Some test cases rely on a second SUT. If the second SUT is
not present, then some steps of these test cases will not be run.

### MS-ADMINS

Two scenarios are designed to verify the server-side, testable
requirements in MS-ADMINS test suite. The following table lists the
scenarios designed in the test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_CreateAndDeleteSite     |     A client tries to create a site or delete a site and tries to get LCID from the server.
  S02\_ErrorConditions         |     Test the negative conditions when the protocol client calls the CreatSite or DeleteSite operations.

### MS-AUTHWS

Four scenarios are designed to verify the server-side, testable
requirements in MS-AUTHWS test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_LoginApplicationUnderFormsAuthentication     |     This scenario is designed to test login application under Forms authentication mode.
  S02\_LoginApplicationUnderNoneAuthentication      |     This scenario is designed to test login application under None authentication mode.
  S03\_LoginApplicationUnderWindowsAuthentication   |     This scenario is designed to test login application under Windows authentication mode.
  S04\_LoginApplicationUnderPassportAuthentication  |     This scenario is designed to test login application under Passport authentication mode.

### MS-COPYS

Two scenarios are designed to verify the server-side, testable
requirements in MS-COPYS test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_CopyIntoItems          |     Copy a file to the destination server, and the destination server is different with the source location.
  S02\_CopyIntoItemsLocal     |     Copy a file to the destination server, and the destination server is same with the source location.

### MS-CPSWS

Five scenarios are designed to verify the server-side, testable
requirements in MS-CPSWS test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_RetrieveTypes                 | This scenario is designed to retrieve a list of all possible claim types, claim value types or entity types from a list of claim providers available to the protocol client.
  S02\_RetrieveProviderHierarchyTree | This scenario is designed to retrieve provider hierarchy trees from a list of claim providers available to the protocol client.
  S03\_RetrieveProviderSchema        | This scenario is designed to retrieve provider schemas.
  S04\_ResolveToEntities             | This scenario is designed to resolve input strings/claims to picker entities using a list of claim providers.
  S05\_SearchForEntities             | This scenario is designed to search for entities on a list of claims providers.

### MS-DWSS

Five scenarios are designed to verify the server-side, testable
requirements in MS-DWSS test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_ManageSites       This scenario is designed to manage Document Workspace sites.
  S02\_ManageData        This scenario is designed to manage data for the Document Workspace site.
  S03\_ManageFolders     This scenario is designed to manage folders in the Document Workspace site.
  S04\_ManageDocuments   This scenario is designed to manage documents in the Document Workspace site.
  S05\_ManageSiteUsers   This scenario is designed to manage site users for the Document Workspace site.

### MS-LISTSWS

Five scenarios are designed to verify the server-side, testable
requirements in MS-LISTSWS test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_OperationOnList        |  Implement the operations on lists and list collection.
  S02\_OperationOnContentType |  Implement the operations on content types and content type XML documents.
  S03\_OperationOnListItem    |  Implement the operations on list items.
  S04\_OperationOnAttachment  |  Implement the operations on attachments.
  S05\_OperationOnFiles       |  Implement the operations on files.

### MS-MEETS

Four scenarios are designed to verify the server-side, testable
requirements in MS-MEETS test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_MeetingWorkspace  | This scenario is to add meeting workspace, set workspace’s title, get workspaces information and delete the meeting workspace.
  S02\_Meeting           | This scenario is to add meeting, update meeting, delete and restore the meeting.
  S03\_MeetingFromICal   | This scenario is to add, update and meeting to a workspace based on a calendar object. Also include set attendee response.
  S04\_RecurringMeeting  | This scenario is to add recurring meeting to a workspace.

### MS-OFFICIALFILE

Four scenarios are designed for this test suite to verify the
server-side, testable requirements in MS-OFFICIALFILE test suite. The
following table lists the scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_GetRoutingDestinationUrlAndSubmitFile  | This scenario is used to test GetRoutingDestinationUrl and SubmitFile operations.
  S02\_GetRoutingInfo                         | This scenario is used to test GetRecordRouting and GetRecordRoutingCollection operations.
  S03\_GetServerInfo                          | This scenario is used to test GetServerInfo operation.
  S04\_GetHoldsInfo                           | This scenario is used to test GetHoldsInfo operation.

### MS-OUTSPS

Three scenarios are designed to verify the server-side, testable
requirements in MS-OUTSPS test suite. The following table lists the
scenarios designed in this test suite

  Scenario  |  Description
:------------ | :-------------
  S01\_OperateAttachment   | The client tries to add an attachment on a list item, perform update/delete operation on the attachment.
  S02\_OperateListItems    | The client tries to add list items on a specified list, and perform update/delete operation on these list items, and sync the list items changes from the protocol SUT.
  S03\_CheckListDefination |  The client tries to get the list definition from the protocol SUT, and verify the field definition of specified list.

### MS-SHDACCWS

One scenario is designed to verify the server-side, testable
requirements in MS-SHDACCWS test suite. The following table lists the
scenarios designed in this test suite.

Scenario  |  Description
:------------ | :-------------
  S01\_VerifyIsSingleClient   | This scenario is used to judge whether it is the only client currently editing a document stored on a collaboration server, or alternately, whether should transition to a shared editing mode.

### MS-SITESS

Seven scenarios are designed to verify the server-side and testable
requirements in MS-SITESS test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_MigrateSite             | This scenario is designed to migrate the content of a site.
  S02\_ManageSubSite           | This scenario is designed to get available site template information and create/delete a subsite.
  S03\_GetUpdatedFormDigest    | This scenario is designed to get a new form digest validation and its expiration time.
  S04\_ExportSolution          | This scenario is designed to export the content related to a site to the solution gallery.
  S05\_ExportWorkflowTemplate  | This scenario is designed to export a workflow template as a site solution to the document library.
  S06\_GetSite                 | This scenario is designed to get information about site collection.
  S07\_HTTPStatusCode          | This scenario is designed to send request with an unauthenticated account in order to trigger an HTTP Status Code fault.

### MS-VERSS

Three scenarios are designed to verify the server-side, testable
requirements in MS-VERSS test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_DeleteVersion    | Get and delete versions for a specified file with valid input parameters.
  S02\_RestoreVersion   | Get and restore versions for a specified file with valid input parameters.
  S03\_ErrorConditions  | Verify various error conditions of the 4 operations.

### MS-VIEWSS

Five scenarios are designed to verify the server-side,testable
requirements in MS-VIEWSS test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_AddDeleteViews                                      | Add and delete a view for a specified list with valid or invalid input parameters. Get the added or deleted view for verification.
  S02\_GetAllViews                                         | Get all of the views of a specified list with valid or invalid input parameters.
  S03\_MaintainViewDefinition                              | Get or update the definition of a view of a specified list with valid or invalid input parameters.
  S04\_MaintainViewProperties                              | Get or update the definition and display properties of a view of a specified list with valid or invalid input parameters.
  S05\_MaintainViewPropertiesWithOpenApplicationExtension  | Get or update the definition and display properties of a view of a specified list with valid or invalid input parameters, with open application extension.

### MS-WDVMODUU

Three scenarios are designed to verify the server-side,testable
requirements in MS-WDVMODUU test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_XVirusInfectedHeader  | This scenario is designed to test the extension header “XVirusInfectedHeader” in the HTTP GET/PUT response
  S02\_IgnoredHeaders        | Send message to the server by HTTP/1.1 PUT/GET/DELETE requests to test the extension headers that are ignored by protocol server in this protocol.
  S03\_PropFindExtension     | This scenario is designed to test the extension properties in the XML body of HTTP PROPFIND method.

### MS-WEBSS

Ten scenarios are designed to verify the server-side,testable
requirements in MS-WEBSS test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_OperationsOnContentType             |Protocol client tries to perform operations associated with content types.
  S02\_OperationsOnContentTypeXmlDocument  |Protocol client tries to perform operations associated with XML document.
  S03\_OperationsOnPage                    |Protocol client tries to perform operations associated with page.
  S04\_OperationsOnFile                    |Protocol client tries to perform operations associated with file.
  S05\_OperationsOnObjectId                |Protocol client tries to perform operations associated with objectId.
  S06\_OperationsOnListTemplates           |Protocol client tries to get the collection of list templates definitions.
  S07\_OperationsOnColumns                 |Protocol client tries to perform operations associated with columns.
  S08\_OperationsOnCSS                     |Protocol client tries to perform operations associated with the customization of the specified CSS.
  S09\_OperationsOnWeb                     |Protocol client tries to perform operations associated with sub-webs, webs and web collection.
  S10\_OperationsOnActivatedFeatures       |Protocol client tries to perform operations associated with activated features.

### MS-WSSREST

Three scenarios are designed to verify the server-side,testable
requirements in MS-WSSREST test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_ManageListItem        | Operate the operations of retrieve, insert, update and delete on List Item.
  S02\_RetrieveCSDLDocument  | Retrieve the conceptual schema definition language (CSDL) document.
  S03\_BatchRequests         | Implement multiple operations in a HTTP Request.

### MS-WWSP

Four scenarios are designed to verify the server-side,testable
requirements in MS-WWSP test suite. The following table lists the
scenarios designed in this test suite.

  Scenario  |  Description
:------------ | :-------------
  S01\_StartWorkflow     | The client tries to start a workflow task for a document item by specified workflow association.
  S02\_GetForItem        | The client tries to get workflow related data from SUT, include workflow association data, workflow task and workflow data.
  S03\_AlterToDo         | The client tries to update a workflow task.
  S04\_ClaimReleaseTask  | The client tries to start a workflow task for a document, claim it and then release it.



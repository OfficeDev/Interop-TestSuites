Exchange EWS Test Suite deployment guide
======================================================================================================
- [Overview](#overview)
- [Prerequisites](#prerequisites)
- [Deploying test suites](#deploying-test-suites)
- [Using test suite directories](#using-test-suite-directories)
- [Configuring test suites](#configuring-test-suites)
- [Running test suites](#running-test-suites)
- [Viewing test suite results, logs, and reports](#viewing-test-suite-results-logs-and-reports)
- [Appendix](#appendix)

Overview
======================================================================================================
Exchange Server EWS Protocol Test Suites are implemented as
synthetic clients running against the server-side implementation of a
given Exchange protocol. They are designed in a client-to-server
relationship and were originally developed for the in-house testing of
Microsoft Open Specifications. Test Suites have been used
extensively in Plugfests and Interoperability Labs to test partner
implementation.

Exchange EWS Test Suite deployment guide introduces the hardware and
software requirements of the test suite client, and the requirements of
the system under test (SUT) if the test suites run against Exchange
Server. The guide also introduces topics on how to deploy, configure and run the
test suites, and view the test suite reports.

Prerequisites
========================================================================================================================================================================================================

This section describes the hardware and software requirements for the
test suites. In an Exchange server environment, the test suite
installation takes place on both the client and server side. The
following information helps test suite users to plan their
deployment.

Hardware requirements
------------------------------------------------------------------------------------------------------------------------------------------------------------------

### System under test

The SUT is the server side of the test suite environment. Exchange
server(s) and Active Directory have defined system requirements which
should be taken into account during deployment. Exchange Server EWS
Protocol Test Suites do not have any additional SUT resource
requirements.

### Test suite client

The test suite client is the client side of the test suite environment.
The following table shows the minimum resource requirements for the test
suite client.

**Test suite client resource requirements**

Component| Test suite client minimum requirement
:------------ | :-------------
**RAM**       |  2GB
**Hard Disk** |  3G of free space
**Processor** |  >= 1GHz

Software requirements
-------------------------------------------------------------------------------------------------------------------

### System under test

This section is only relevant when running the test suites against the
following versions of Exchange Server:

-   Microsoft Exchange Server 2007 Service Pack 3 (SP3)
-   Microsoft Exchange Server 2010 Service Pack 3 (SP3)
-   Microsoft Exchange Server 2013 Service Pack 1 (SP1)
-	Microsoft Exchange Server 2016
-	Microsoft Exchange Server 2019

The following table describes the required server roles for a
test suite deployment with Microsoft implementation. Installing
Exchange Server on a domain controller (DC) is not recommended.

**Required SUT roles**

Role | Description
:------------ | :-------------
**Active Directory Domain Controller (AD DC)** |  Active Directory Domain Controller provides secure data for users and computers. An AD DC can coexist with Exchange Server. A typical test configuration has an AD DC and Exchange Server installed on separate machines.
**Exchange Server (SUT)**  |  The Exchange Server in the topology.

The following diagram is an example of what a typical Exchange test suite
environment may look like. This example uses an IPv4, but IPv6 is also
supported by the test suites.
![alt tag](./Doc-Images/EWS_RequiredSUTRoles.png)

### Test suite client

This section describes the prerequisite software for installing Exchange Server EWS Protocol Test Suites on the test suite client. The following
table outlines the software dependencies for the test suite client.

**Test suite client software dependencies**

| Operating systems |
| :------------|
| Windows 7 x64 Service Pack 1 and above |
|Windows 8 x64 and above|
|Windows 2008 R2 x64 Service Pack 1 and above|

| Software |
| :------------|
| Microsoft Visual Studio 2013 Professional |
| Microsoft Protocol Test Framework 1.0.2220.0 and above|

Deploying test suites
=======================================================================================================================

This section describes the deployment of Exchange Server EWS Protocol
Test Suites on the test suite client and the SUT. Exchange Server
EWS Protocol Test Suites are packaged in a .zip file, available
at [Microsoft Connect](http://go.microsoft.com/fwlink/?LinkId=516921).
Once you've downloaded the test suites, perform the following
steps to successfully configure the test suites.

1.  Extract the **Exchange Server EWS Protocol Test Suites** folder from the zip file to a
    directory of your choice on the test suite client.

2.  Copy the **SUT** folder under **…\\Exchange Server EWS Protocol Test
    Suites\\Setup** to a directory of your choice on the SUT. The SUT
    configuration scripts are the only requirement for the SUT. The
    scripts facilitate the SUT configuration process in the **ExchangeServerEWSProtocolTestSuites.zip** file.

Using test suite directories
================================================================================================================================================================================================================================================================================================================================================================================================================

This section shows the folder structures in the **ExchangeServerEWSProtocolTestSuites.zip** file.

**ExchangeServerEWSProtocolTestSuites.zip file contents**

Folder/file  | Description
:------------   | :-------------
**EULA.rtf**    | End-User License Agreement.
**ReadMe.txt**  | A doc on deployment and prerequisite software.
**Exchange Server EWS Protocol Test Suites**  |--
**- Docs**      |A folder with documents of all protocol test suites.
**- ExchangeEWSTestSuiteDeploymentGuide.docx**  |  A doc on the protocol test suite deployment.
**-ExchangeEWSTestSuiteSpecification.docx**     |  A doc on the test suite configuration details, architecture, adapters and test case details.												   
**+ MS-XXXX**                                   |  The MS-XXXX help documentation.
**- \[MS-XXXX\].pdf**                           |  The protocol technical specification.
**- MS-XXXX \_SUTControlAdapter.chm**           |  A help doc on the SUT control adapter class library such as declaration syntax and their description.
**- MS-XXXX \_RequirementSpecification.xlsx**   |  A spreadsheet that outlines all requirements that are associated with the technical specification.
**- Setup**                                     |  A folder with configuration scripts.
**- Test Suite Client**                         |  A folder with the configuration script to configure the test suite client.
**- ExchangeClientConfiguration.cmd**           |  A command file that runs the ExchangeClientConfiguration.ps1 file to configure the properties for the protocol test suites.
**- ExchangeClientConfiguration.ps1**           |  A configuration script that will be triggered by ExchangeClientConfiguration.cmd.
**- SUT**                                       |  A folder with the configuration script to configure Exchange Server.
**- ExchangeSUTConfiguration.cmd**              |  A command file that runs the ExchangeSUTConfiguration.ps1 file to create resources and configure settings on the SUT.
**- ExchangeSUTConfiguration.ps1**              |  A configuration script that will be triggered by ExchangeSUTConfiguration.cmd.
**- Common**                                    |  A folder with common configuration scripts.
**- CommonConfiguration.ps1**                   |  A configuration script to configure the common information of the server and the test suite client.
**- ExchangeCommonConfiguration.ps1**           |  A configuration script to configure the common information of Exchange Server.
**- ExchangeTestSuite.config**                  |  A configuration file that contains primary SUT configuration resources of all protocol test suites.
**- Source**                                    |  A folder with Microsoft Visual Studio solutions that contain the source code for the test suites.
**- Common**                                    |  A folder with Microsoft Visual Studio projects that contains the common source code for the test suites.
**- ExchangeCommonConfiguration.deployment.ptfconfig** |  The common configuration file.
**- ExchangeEWSProtocolTestSuites.sln**                |  A Visual Studio solution with projects that encapsulate the protocol test suites source code.
**- ExchangeServerEWSProtocolTestSuites.runsettings**  |  A configuration file for the unit test.
**- ExchangeServerEWSProtocolTestSuites.testsettings** |  A configuration file for running test cases.
**- MS-XXXX**                                          |  A folder for the MS-XXXX test suite source code.
**- MS-XXXX.sln**                                      |  A Microsoft Visual Studio solution that contains projects of the MS-XXXX test suite.
**- MS-XXXX.runsettings**                              |  A configuration file for the MS-XXXX unit test.
**- MS-XXXX.testsettings**                             |  A configuration file for running MS-XXXX test cases.
**+ Adapter**                                          |  The Adapter test suite code.
**+ TestSuite**                                        |  The test suite code.
**- Scripts**                                          |  Exchange Server EWS Test Suites can be run using Visual Studio or batch scripts. The Scripts folder contains a collection of command files that allows users to run specific test cases in the test suite, or the entire test suite.
**- RunAllExchangeEWSTestCases.cmd**                   |  A script that can be used to run all test cases in the package.
**- MS-XXXX**                                          |  A folder containing scripts that belong to the MS-XXXX test suite.
**- RunAllMSXXXXTestCases.cmd**                        |  A script that can be used to run all test cases of MS-XXXX.
**- RunMSXXXX\_SXX\_TCXX\_Name.cmd**                   |  A script that can be used to run a single test case of MS-XXXX.

Configuring test suites
================================================================================================================================================================================================================================================================================================================================================================================================================

This section provides the guidance on configuring Exchange
Server EWS Protocol Test Suites on the SUT and the test suite client.
The configuration should be done in this order: configure the SUT, and
then configure the test suite client.

For the configuration script, the exit code definition is as follows:

1.  A normal termination will set the exit code to 0.

2.  An uncaught THROW will set the exit code to 1.

3.  Script execution warning and issues will set the exit code to 2.

4.  Exit code is set to the actual error code for other issues.

Configuring the SUT
----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

You can configure the SUT using automated scripts, as described in [Configuring the SUT using the setup configuration script](#configuring-the-sut-using-the-setup-configuration-script); or configure the SUT manually, as described in [Configuring the SUT manually](#configuring-the-sut-manually).

**Note**   The scripts should be run by a user who has domain
administrator rights with a mailbox on the SUT.

### SUT resource requirements

Each test suite in the Exchange Server EWS Protocol Test
Suites package may require varying levels of resources on the SUT. The
following table outlines these resources for each test suite. The SUT
configuration scripts will automatically create all the required
resources for the Microsoft server implementation. To configure the SUT
manually, see [Configuring the SUT manually](#configuring-the-sut-manually).

The client configuration script follows the naming convention shown in
the following table. If a change to the resource name is required, the
corresponding change to the resource name defined in the client
configuration script will be required.

**Exchange server resources**

|Test suite | Resource type |  Resource name |  Notes|
|:------------ | :------------- | :------------- | :-------------|
|**All**         |   --           |          --    |     --                         
|**MS-OXWSATT**  |  Mailbox |  MSOXWSATT\_User01   | Mailbox type user|  
|**MS-OXWSBTRF** |  Mailbox |  MSOXWSBTRF\_User01  | Mailbox type user|  
|**MS-OXWSCONT** |  Mailbox |  MSOXWSCONT\_User01  | Mailbox type user|  
|**MS-OXWSCORE** |  Mailbox |  MSOXWSCORE\_User01  | Mailbox type user|  
|                |  Mailbox |  MSOXWSCORE\_User02  | Mailbox type user|  
|                |  Public Folder Mailbox |  MSOXWSCORE\_PublicFolderMailbox  | Public Folder Mailbox created for the public folder of the organization configuration of Exchange 2013.|  
|                |  Public Folder |  MSOXWSCORE\_PublicFolder      |      
|**MS-OXWSFOLD** |  Mailbox     |  MSOXWSFOLD\_User01   | Mailbox type user|  
|                |  Mailbox       |  MSOXWSFOLD\_User02   | Mailbox type user|  
|                |  ManagedFolder |  MSOXWSFOLD\_ManagedFolder1  | Managed folder created directly in the root path of Outlook|  
|                |  ManagedFolder |  MSOXWSFOLD\_ManagedFolder2  | Managed folder created directly in the root path of Outlook|  
|                |  Public Folder Database  | PublicFolderDatabase     | Public Folder Database created for the mailbox of the organization configuration of Exchange 2010 and for the server configuration of Exchange 2007.|  
| **MS-OXWSMSG**|  Mailbox  |  MSOXWSMSG\_User01  |  Mailbox type user|  
|               |  Mailbox  |  MSOXWSMSG\_User02  |  Mailbox type user|  
|               |  Mailbox  |  MSOXWSMSG\_User03  |  Mailbox type user|  
|               |  Mailbox  |  MSOXWSMSG\_Room01  |  Mailbox type room|  
|  **MS-OXWSMTGS** |  Mailbox|  MSOXWSMTGS\_User01 |  Mailbox type user|  
|                |  Mailbox  |  MSOXWSMTGS\_User02 |  Mailbox type user|  
|                |  Mailbox  |  MSOXWSMTGS\_Room01 |  Mailbox type room|
|                |  Mailbox  |  MSOXWSMTGS\_User03 |  Mailbox type user|  
|  **MS-OXWSSYNC** |  Mailbox|  MSOXWSSYNC\_User01 |  Mailbox type user|  
|                |  Mailbox  |  MSOXWSSYNC\_User02 |  Mailbox type user|  
|  **MS-OXWSTASK** |  Mailbox|  MSOXWSTASK\_User01 |  Mailbox type user|  

### Configuring the SUT using the setup configuration script <a name="configuring-the-sut-using-the-setup-configuration-script"></a>

The setup configuration script is only used for configuring the SUT on
the Windows platform.

To configure SUT using the setup configuration script, navigate to the
**SUT** folder, right-click **ExchangeSUTConfiguration.cmd** and select
**Run as administrator**.

### Configuring the SUT manually <a name="configuring-the-sut-manually" ></a>

If the SUT is non-Microsoft implementation of Exchange Server, you
will not be able to run the setup configuration script. The following
steps explain what needs to be created or configured on the SUT to run the test suites.

1.  Create the following mailbox users:

	MSOXWSATT\_User01, MSOXWSBTRF\_User01, MSOXWSCONT\_User01,
	MSOXWSCORE\_User01, MSOXWSCORE\_User02, MSOXWSFOLD\_User01,
	MSOXWSFOLD\_User02, MSOXWSMSG\_User01, MSOXWSMSG\_User02,
	MSOXWSMSG\_User03, MSOXWSMSG\_Room01, MSOXWSMTGS\_User01,
	MSOXWSMTGS\_User02, MSOXWSMTGS\_Room01, MSOXWSMTGS\_User03, MSOXWSSYNC\_User01,
	MSOXWSSYNC\_User02, MSOXWSTASK\_User01

1.  Configure Secure Sockets Layer (SSL) as **not required** and set to
    ignore client certificates on the website which contains the
    application that implements the EWS protocols.

2.  Assign the **ApplicationImpersonation** role to the following
    mailbox users.

	**Note** This role enables applications to impersonate users in an
	organization to perform a task on behalf of users.

	MS-OXWSATT\_User01, MS-OXWSBTRF\_User01, MSOXWSCORE\_User01,
	MSOXWSFOLD\_User01, and MSOXWSSYNC\_User01

1.  Create the following managed folders in Active Directory:

	MSOXWSFOLD\_ManagedFolder1 and MSOXWSFOLD\_ManagedFolder2

1.  Create a public folder database.

2.  Grant permissions to the mailbox user MSOXWSFOLD\_User01 to manage
    the public folders.

3.  Create a public folder MSOXWSCORE\_PublicFolderMailbox.

4.  Grant permissions to the mailbox user MSOXWSCORE\_User01 to manager
    the public folder MSOXWSCORE\_PublicFolder.

Configuring the test suite client
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

The test suite client is managed through a common configuration file,
two test-suite specific configuration files, and three SHOULD/MAY
configuration files that all have a “.ptfconfig” extension. These
configuration files can be modified directly. The common configuration
file and the test-suite specific configuration files can also be
modified through a script.

### Common configuration file

The common configuration file contains configurable properties common to
all Exchange Server EWS Protocol Test Suites. This file must be modified
to match the characteristics of the environment where the test suites
are installed.

Configuration file| Description
:------------ | :-------------
**ExchangeCommonConfiguration.deployment.ptfconfig** |   The deployment configuration file provides the common environmental details for the test suites.

### Test-suite specific configuration files

In addition to the common configuration file, each individual test suite
has the following two configuration files for test-suite specific
modification.

**Test-suite specific configuration files**

Configuration file  | Description
:------------ | :-------------
**MS-XXXX\_TestSuite.deployment.ptfconfig** |The deployment configuration file provides the environmental details that are specific to the test suite. The configuration file allows for the test- suite specific customization.
**MS-XXXX\_TestSuite.ptfconfig**            |The test suite configuration file contains details that specify the behavior of the test suite operation.

Both files are in the TestSuite folder in each test suite directory.

If you need to modify the common configuration values for a specific
test suite, you must copy the common properties to the
**MS-XXXX\_TestSuite.deployment.ptfconfig** file and change the values
of the properties. The specific configuration file will take precedence
over the common configuration file when the same property exists in both
places.

#### Set the test suite to interactive mode <a name="set-the-test-suite-to-interactive mode"></a>

If the SUT is non-Microsoft implementation of Exchange Server, it is
recommended that you further configure the test suite by setting the
test suite to interactive mode. Interactive mode enables the test suite
to function in a manual way, enabling you to perform setup, teardown,
and other tasks in a step-by-step approach. To enable interactive mode
for a specific test suite, do the following:

1.  Browse to the **MS-XXXX\_TestSuite.ptfconfig** configuration file in **\\Source\\MS-XXXX\\TestSuite\\**.

1.  Set the type value of Adapter property to **interactive** for the
    SUT control adapter\*\*.

**Interactive mode values**

Property name | Default value\*|Optional value  |  Description|
:------------ | :------------- | :------------- | :-------------
Adapter    |     managed or powershell|   interactive\*\* |  **managed**: The SUT control adapter is implemented in C\# managed code.
 ||| **powershell**: The SUT control adapter is implemented through Windows PowerShell.
 ||| **interactive**: Interactive adapters are used for manually configuring the server. Interactive adapter displays a dialog box to perform a manual test each time when one of its methods is called. The dialog box will show the method name, parameter names, and values\*\*\*

\*The Adapter property value is set to either managed or powershell
depending on whether the SUT control adapter was implemented in managed
C\# code or through PowerShell.

\*\*When changing from managed mode to interactive mode, the
“adaptertype” attribute must be deleted to avoid a runtime error. When
changing from powershell mode to interactive mode, an additional step is
required—delete the “scriptdir” attribute to avoid a runtime error.

\*\*\*When the manual operation completes successfully, enter the
return values (if any) in **Action Results** and click **Succeed** in
the dialog box. When the manual operation is unable to complete, enter
the error messages in the **Failure Message** text box and click
**Fail** to terminate the test. In this case, the test will be treated
as “Inconclusive”.

Further customization can be done by creating your own SUT control
adapter that matches the server implementation. For more information
about how to create a SUT control adapter, see the [Protocol Test
Framework (PTF) user documentation](https://github.com/Microsoft/ProtocolTestFramework).

#### Configure TSAP broadcast

Test Session Announcement Protocol (TSAP) is used by PTF to broadcast
test information when the test suite is running. TSAP broadcast helps with mapping test cases to captured frames.

By default, TSAP packets are broadcasted in the network. The user can
change a TSAP broadcast by adding an entry “BeaconLogTargetServer” to
TestSuite.deployment.ptfconfig to target TSAP for the specified
machine.

To change the TSAP packet broadcast, do the following:

1.  Browse to the **MS-XXXX\_TestSuite.deployment.ptfconfig**
    configuration file in the **\\Source\\MS-XXXX\\TestSuite\\** folder.

2.  Add a property “BeaconLogTargetServer” along with the value of the
    specified machine name.

	For example: &lt;Property name="BeaconLogTargetServer" value="dc01"/&gt;

### SHOULD/MAY configuration files

The test suite has three SHOULD/MAY configuration files that are
specific to all supported versions of the SUT. Each SHOULD/MAY
requirement has an associated parameter with a value of either “true”
or “false” corresponding to the server version that is supported. The value of “true”
means that the requirement must be validated, whereas “false” means
that the requirement must not be validated.

If the SUT is non-Microsoft implementation of Exchange Server,
configure the properties in the configuration file for the Exchange
Server to be the closest match to the SUT implementation.

**SHOULD/MAY configuration files**

Configuration file | Description
:------------ | :-------------
**MS-XXXX\_ExchangeServer2007\_SHOULDMAY.deployment.ptfconfig** | Provides the configuration properties for SHOULD and MAY requirements supported by Microsoft Exchange Server 2007 Service Pack 3 (SP3).
**MS-XXXX\_ExchangeServer2010\_SHOULDMAY.deployment.ptfconfig** | Provides the configuration properties for SHOULD and MAY requirements supported by Microsoft Exchange Server 2010 Service Pack 3 (SP3).
**MS-XXXX\_ExchangeServer2013\_SHOULDMAY.deployment.ptfconfig** | Provides the configuration properties for SHOULD and MAY requirements supported by Microsoft Exchange Server 2013 Service Pack 1 (SP1).
**MS-XXXX\_ExchangeServer2016\_SHOULDMAY.deployment.ptfconfig** | Provides the configuration properties for SHOULD and MAY requirements supported by Microsoft Exchange Server 2016.
**MS-XXXX\_ExchangeServer2019\_SHOULDMAY.deployment.ptfconfig** | Provides the configuration properties for SHOULD and MAY requirements supported by Microsoft Exchange Server 2019.

### Configuring the test suite client using the setup configuration script

**Note** The setup configuration script is only implemented for configuring the test
suite client on the Windows platform.

To configure the test suite using the setup configuration script,
navigate to the **Setup\\Test Suite Client**\\ folder, right-click
**ExchangeClientConfiguration.cmd** and select **Run as administrator**.

### Configuring the test suite client manually

If you didn’t use the setup configuration script to configure the test
suite client as described in the previous section, follow the steps
below to update configuration files and configure the test suite client.

1.  Update the property value in the common configuration file and the
    test-suite specific configuration files according to the comment of
    the property.

Running test suites
========================================================================================================================================================================================================

Once the required software is installed and both the SUT and test suite client are configured appropriately, the test suite is ready to run. The test
suite can run only on the test suite client and can be initiated in one
of the following two ways: Visual Studio or batch scripts.

Microsoft Visual Studio
---------------------------------------------------------------------------------------------------------------------

A Microsoft Visual Studio solution file
**ExchangeServerEWSProtocolTestSuites.sln** is provided in the
**Source** folder. You can run a single or multiple test cases in Visual
Studio.


1.  Open **ExchangeServerEWSProtocolTestSuites.sln** in Visual Studio.                                                                       
![alt tag](./Doc-Images/EWS_RunningTS1.png)                                                                                                                                         

  -------------------------------------------------------------------------------------------------------------------------------------------- --
2.  In the **Solution Explorer** pane, right-click **Solution ‘ExchangeServerEWSProtocolTestSuites’** and then click **Rebuild Solution**.   
![alt tag](./Doc-Images/EWS_RunningTS2.png)                                                                                                                                         
   -------------------------------------------------------------------------------------------------------------------------------------------- --                                                                                                                                            

3.  Open **Test Explorer**. On the ribbon click **TEST**, then click **Windows**, and finally click **Test Explorer**.                       
![alt tag](./Doc-Images/EWS_RunningTS3.png)                                                                                                                                      
   -------------------------------------------------------------------------------------------------------------------------------------------- --                                                                                                                                            

4.  Select the test case to run. Right-click and then select **Run Selected Tests**.
![alt tag](./Doc-Images/EWS_RunningTS4.png)
  -----------------------------------------------------------------------------------------------------------------------------------------------

A Visual Studio solution file
**MS-XXXX.sln** is provided in each test suite folder.

1.  Select the test suite you would like to run. Let’s take MS-OXWSCORE as an example here, so browse to the **Source\\MS-OXWSCORE** directory.

  ------------------------------------------------------------------------------------------------------------------------------------------------- -----------------------------------------------------------------------------------
2.  Open **MS-OXWSCORE.sln** in Microsoft Visual Studio.

![alt tag](./Doc-Images/EWS_RunningTS5.png)

  -------------------------------------------------------------------------------------------------------------------------------------------- --

3.  In the **Solution Explorer** pane, right-click **Solution ‘MS-OXWSCORE’**, and then click **Rebuild Solution**.

![alt tag](./Doc-Images/EWS_RunningTS6.png)

  -------------------------------------------------------------------------------------------------------------------------------------------- --

4.  Open **Test Explorer**. On the ribbon click **TEST**, then click **Windows**, and finally click **Test Explorer**.

![alt tag](./Doc-Images/EWS_RunningTS7.png)

  -------------------------------------------------------------------------------------------------------------------------------------------- --

5.  Select the test case to run. Right-click and then select **Run Selected Tests**.

![alt tag](./Doc-Images/EWS_RunningTS8.png)


Batch scripts
------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

Exchange Server EWS Protocol Test Suites are installed with a collection
of scripts that enable a user to run individual test cases
(RunMSXXXX\_SXX\_TCXX\_Name.cmd) or all test cases in a test suite
(RunAllMSXXXXTestCases.cmd), or all test cases of Exchange Server EWS
Protocol Test Suites at once (RunAllExchangeEWSTestCases.cmd). These
scripts can be found in the **\\Source\\Scripts** directory.

**Note**  These scripts depend on having the compiled binaries in the bin folder.

Batch script | Script description
:------------ | :-------------
**RunAllExchangeEWSTestCases.cmd**  |  Runs all the test cases in Exchange Server EWS Protocol Test Suites.
**RunAllMSXXXXTestCases.cmd**       |  Runs all MS-XXXX test cases.
**RunMSXXXX\_SXX\_TCXX\_Name.cmd**  |  Runs a specific test case in the test suite.

Viewing test suite results, logs, and reports
=====================================================================================================================================

The test suites provide detailed reporting in a variety of formats that
enable users to quickly debug failures.

Test suite configuration logs
---------------------------------------------------------------------------------------------------------------------------

The configuration logs show whether or not each
configuration step succeeds and detailed information on errors if the
configuration step fails.

### SUT configuration logs

The SUT configuration scripts create a directory named **SetupLogs**
under **…\\Setup\\SUT\\** at runtime. The SUT configuration scripts save
the logs as “ExchangeSUTConfiguration.ps1.debug.log” and “ExchangeSUTConfiguration.ps1.log”.

### Test suite client configuration logs

The configuration scripts create a directory named **SetupLogs** under
**…\\Setup\\Test Suite Client\\** at runtime. The test suite client
configuration scripts save the logs as “ExchangeClientConfiguration.ps1.debug.log” and “ExchangeClientConfiguration.ps1.log”.

Test suite reports
------------------

### Microsoft Visual Studio

Reports are created only after the package level solution or an
individual test suite solution has run successfully in Visual Studio.

-   Reporting information for
    **ExchangeServerEWSProtocolTestSuites.sln** is saved in
    **…\\Source\\TestResults**.

-   Reporting information for an individual test suite **MS-XXXX.sln**
    is saved in **…\\Source\\MS-XXXX\\TestResults**.

### Batch scripts

If Exchange Server EWS Protocol Test Suites are run by the
RunAllExchangeEWSTestCases.cmd batch file, the reporting information is
saved in **…\\Source\\Scripts\\TestResults**.

If the test suite is run by the batch file RunAllMSXXXXTestCases.cmd or
RunMSXXXX\_SXX\_TCXX\_Name.cmd, the reporting information is saved in
**…\\Source\\Scripts\\MS-XXXX\\TestResults.**

By default, a .trx file containing the pass/fail information of the run
is created in the TestResults folder along with an associated directory
named **user\_MACHINENAME DateTimeStamp** that contains a log file in an
XML format and an HTML report.

Appendix
======================================================================================================
For more information, see the following:

References | Description
:------------ | :-------------
<dochelp@microsoft.com>   |  The alias for Interoperability documentation help, which provides support for Open Specifications and protocol test suites.
[Open Specifications Forums](http://go.microsoft.com/fwlink/?LinkId=111125)  |   The Microsoft Customer Support Services forums, the actively monitored forums that provides support for Open Specifications and protocol test suites. |
[Open Specifications Developer Center](http://go.microsoft.com/fwlink/?LinkId=254469)   |  The Open Specifications home page on MSDN.
[Open Specifications](http://go.microsoft.com/fwlink/?LinkId=179743)                    |   The Open Specifications documentation on MSDN.
[Exchange Products and Technologies Protocols](http://go.microsoft.com/fwlink/?LinkId=119904)    |The Exchange Server Open Specifications documentation on MSDN.
[RFC2119](http://go.microsoft.com/fwlink/?LinkId=117453)                                         |The normative language reference.
[Exchange Server 2016/2019 deployment](https://learn.microsoft.com/en-us/exchange/plan-and-deploy/plan-and-deploy?view=exchserver-2019#deploy-exchange-2016-or-exchange-2019)              | The Exchange Server 2016/2019 planning and deployment on TechNet.
[Exchange Server 2013 deployment](http://go.microsoft.com/fwlink/?LinkID=266569)                 |The Exchange Server 2013 planning and deployment on TechNet.
[Exchange Server 2010 deployment](http://go.microsoft.com/fwlink/?LinkID=517397)                 |The Exchange Server 2010 planning and deployment on TechNet.
[Exchange Server 2007 deployment](http://go.microsoft.com/fwlink/?LinkID=512508)                 |The Exchange Server 2007 deployment on TechNet.

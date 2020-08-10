namespace Microsoft.Protocols.TestSuites.MS_OUTSPS
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Xml;
    using System.Xml.Schema;
    using Microsoft.Protocols.TestSuites.Common;
    using Microsoft.Protocols.TestTools;
    using Microsoft.VisualStudio.TestTools.UnitTesting;

    /// <summary>
    /// A test class contains test cases of S03 scenario.
    /// </summary>
    [TestClass]
    public class S03_CheckListDefination : TestSuiteBase
    {
        #region Additional test attributes, Initialization and clean up

        /// <summary>
        /// Test class level initialization method
        /// </summary>
        /// <param name="testContext">The context of the test suite.</param>
        [ClassInitialize]
        public static new void ClassInitialize(TestContext testContext)
        {
            TestSuiteBase.ClassInitialize(testContext);
        }

        /// <summary>
        /// Test class level clean up method
        /// </summary>
        [ClassCleanup]
        public static new void ClassCleanup()
        {
            TestSuiteBase.ClassCleanup();
        }

        #endregion

        #region Test cases

        #region MSOUTSPS_S03_TC01_VerifyAppointmentsList

        /// <summary>
        /// This test case is used to verify definition of Event list which contains Appointments list items.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S03_TC01_VerifyAppointmentsList()
        {
            // Add a Events list.
            string listId = this.AddListToSUT(TemplateType.Events);

            // Call GetList operation
            GetListResponseGetListResult getListResult = OutspsAdapter.GetList(listId);

            #region Verify fields definition

            // Verify Description field's id and type.
            bool isVerifyR594 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Description",
                                            "{9da97a8a-1da5-4a77-98d3-4bc10456e700}",
                                            "Note");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R594
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR594,
                                            594,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Description[ Field.ID:]{9da97a8a-1da5-4a77-98d3-4bc10456e700}[Field.Type:]Note.");

            // Verify Duration field's id and type.
            bool isVerifyR595 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Duration",
                                            "{4d54445d-1c84-4a6d-b8db-a51ded4e1acc}",
                                            "Integer");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R595
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR595,
                                            595,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Duration[ Field.ID:]{4d54445d-1c84-4a6d-b8db-a51ded4e1acc}[Field.Type:]Integer.");

            // Verify Editor field's id and type.
            bool isVerifyR596 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Editor",
                                            "{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}",
                                            "User");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R596
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR596,
                                            596,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Editor[ Field.ID:]{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}[Field.Type:]User.");

            // Verify EndDate field's id and type.
            bool isVerifyR597 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "EndDate",
                                            "{2684f9f2-54be-429f-ba06-76754fc056bf}",
                                            "DateTime");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R597
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR597,
                                            597,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]EndDate[ Field.ID:]{2684f9f2-54be-429f-ba06-76754fc056bf}[Field.Type:]DateTime.");

            // Verify EventDate field's id and type.
            bool isVerifyR598 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "EventDate",
                                            "{64cd368d-2f95-4bfc-a1f9-8d4324ecb007}",
                                            "DateTime");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R598
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR598,
                                            598,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]EventDate[ Field.ID:]{64cd368d-2f95-4bfc-a1f9-8d4324ecb007}[Field.Type:]DateTime.");

            // Verify EventType field's id and type.
            bool isVerifyR599 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "EventType",
                                            "{5d1d4e76-091a-4e03-ae83-6a59847731c0}",
                                            "Integer");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R599
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR599,
                                            599,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]EventType[ Field.ID:]{5d1d4e76-091a-4e03-ae83-6a59847731c0}[Field.Type:]Integer.");

            // Verify fAllDayEvent field's id and type.
            bool isVerifyR600 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "fAllDayEvent",
                                            "{7d95d1f4-f5fd-4a70-90cd-b35abc9b5bc8}",
                                            "AllDayEvent");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R600
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR600,
                                            600,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]fAllDayEvent[ Field.ID:]{7d95d1f4-f5fd-4a70-90cd-b35abc9b5bc8}[Field.Type:]AllDayEvent.");

            // Verify fRecurrence field's id and type.
            bool isVerifyR603 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "fRecurrence",
                                            "{f2e63656-135e-4f1c-8fc2-ccbe74071901}",
                                            "Recurrence");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R603
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR603,
                                            603,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]fRecurrence[ Field.ID:]{f2e63656-135e-4f1c-8fc2-ccbe74071901}[Field.Type:]Recurrence.");

            // Verify Location field's id and type.
            bool isVerifyR606 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Location",
                                            "{288f5f32-8462-4175-8f09-dd7ba29359a9}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R606
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR606,
                                            606,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Location[ Field.ID:]{288f5f32-8462-4175-8f09-dd7ba29359a9}[Field.Type:]Text.");

            // Verify MasterSeriesItemID field's id and type.
            bool isVerifyR607 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "MasterSeriesItemID",
                                            "{9b2bed84-7769-40e3-9b1d-7954a4053834}",
                                            "Integer");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R607
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR607,
                                            607,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]MasterSeriesItemID[ Field.ID:]{9b2bed84-7769-40e3-9b1d-7954a4053834}[Field.Type:]Integer.");

            // Verify RecurrenceData field's id and type.
            bool isVerifyR610 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "RecurrenceData",
                                            "{d12572d0-0a1e-4438-89b5-4d0430be7603}",
                                            "Note");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R610
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR610,
                                            610,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]RecurrenceData[ Field.ID:]{d12572d0-0a1e-4438-89b5-4d0430be7603}[Field.Type:]Note.");

            // Verify RecurrenceID field's id and type.
            bool isVerifyR611 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "RecurrenceID",
                                            "{dfcc8fff-7c4c-45d6-94ed-14ce0719efef}",
                                            "DateTime");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R611
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR611,
                                            611,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]RecurrenceID[ Field.ID:]{dfcc8fff-7c4c-45d6-94ed-14ce0719efef}[Field.Type:]DateTime.");

            // Verify TimeZone field's id and type.
            bool isVerifyR612 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "TimeZone",
                                            "{6cc1c612-748a-48d8-88f2-944f477f301b}",
                                            "Integer");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R612
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR612,
                                            612,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]TimeZone[ Field.ID:]{6cc1c612-748a-48d8-88f2-944f477f301b}[Field.Type:]Integer.");

            // Verify Title field's id and type.
            bool isVerifyR613 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Title",
                                            "{fa564e0f-0c70-4ab9-b863-0177e6ddd247}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R613
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR613,
                                            613,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Title[ Field.ID:]{fa564e0f-0c70-4ab9-b863-0177e6ddd247}[Field.Type:]Text.");

            // Verify UID field's id and type.
            bool isVerifyR614 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "UID",
                                            "{63055d04-01b5-48f3-9e1e-e564e7c6b23b}",
                                            "GUID");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R614
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR614,
                                            614,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]UID[ Field.ID:]{63055d04-01b5-48f3-9e1e-e564e7c6b23b}[Field.Type:]GUID.");

            // Verify XMLTZone field's id and type.
            bool isVerifyR615 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "XMLTZone",
                                            "{c4b72ed6-45aa-4422-bff1-2b6750d30819}",
                                            "Note");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R615
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR615,
                                            615,
                                            @"[In Appointment-Specific Schema][One of the appointment properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]XMLTZone[ Field.ID:]{c4b72ed6-45aa-4422-bff1-2b6750d30819}[Field.Type:]Note.");

            #endregion Verify fields definition

            // If above verification passes, then capture R1066
            // Verify MS-OUTSPS requirement: MS-OUTSPS_R1066
            this.Site.CaptureRequirement(
                            1066,
                            @"[In Message Processing Events and Sequencing Rules][The operation]GetList <2> Gets information about a list.");
        }

        #endregion

        #region MSOUTSPS_S03_TC02_VerifyContactsList

        /// <summary>
        /// This test case is used to verify definition of Contacts list which contains contact list.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S03_TC02_VerifyContactsList()
        {
            // Add a Contacts list.
            string listId = this.AddListToSUT(TemplateType.Contacts);

            // Call GetList operation
            GetListResponseGetListResult getListResult = OutspsAdapter.GetList(listId);

            #region Verify fields

            // Verify CellPhone field's id and type.
            bool isVerifyR571 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "CellPhone",
                                            "{2a464df1-44c1-4851-949d-fcd270f0ccf2}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R571
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR571,
                                            571,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]CellPhone[ Field.ID:]{2a464df1-44c1-4851-949d-fcd270f0ccf2}[Field.Type:]Text.");

            // Verify Comments field's id and type.
            bool isVerifyR566 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Comments",
                                            "{9da97a8a-1da5-4a77-98d3-4bc10456e700}",
                                            "Note");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R566
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR566,
                                            566,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Comments[ Field.ID:]{9da97a8a-1da5-4a77-98d3-4bc10456e700}[Field.Type:]Note.");

            // Verify Company field's id and type.
            bool isVerifyR565 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Company",
                                            "{038d1503-4629-40f6-adaf-b47d1ab2d4fe}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R565
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR565,
                                            565,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Company[ Field.ID:]{038d1503-4629-40f6-adaf-b47d1ab2d4fe}[Field.Type:]Text.");

            // Verify CompanyPhonetic field's id and type.
            bool isVerifyR563 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "CompanyPhonetic",
                                            "{034aae88-6e9a-4e41-bc8a-09b6c15fcdf4}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R563
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR563,
                                            563,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]CompanyPhonetic[ Field.ID:]{034aae88-6e9a-4e41-bc8a-09b6c15fcdf4}[Field.Type:]Text.");

            // Verify Editor field's id and type.
            bool isVerifyR556 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Editor",
                                            "{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}",
                                            "User");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R556
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR556,
                                            556,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Editor[ Field.ID:]{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}[Field.Type:]User.");

            // Verify Email field's id and type.
            bool isVerifyR555 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Email",
                                            "{fce16b4c-fe53-4793-aaab-b4892e736d15}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R555
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR555,
                                            555,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Email[ Field.ID:]{fce16b4c-fe53-4793-aaab-b4892e736d15}[Field.Type:]Text.");

            // Verify FirstName field's id and type.
            bool isVerifyR542 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "FirstName",
                                            "{4a722dd4-d406-4356-93f9-2550b8f50dd0}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R542
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR542,
                                            542,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]FirstName[ Field.ID:]{4a722dd4-d406-4356-93f9-2550b8f50dd0}[Field.Type:]Text.");

            // Verify FirstNamePhonetic field's id and type.
            bool isVerifyR367 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "FirstNamePhonetic",
                                            "{ea8f7ca9-2a0e-4a89-b8bf-c51a6af62c73}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R367
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR367,
                                            367,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field.Field.Name: ]FirstNamePhonetic[ Field.ID:]{ea8f7ca9-2a0e-4a89-b8bf-c51a6af62c73}[Field.Type:]Text.");

            // Verify FullName field's id and type.
            bool isVerifyR539 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "FullName",
                                            "{475c2610-c157-4b91-9e2d-6855031b3538}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R539
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR539,
                                            539,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]FullName[ Field.ID:]{475c2610-c157-4b91-9e2d-6855031b3538}[Field.Type:]Text.");

            // Verify HomePhone field's id and type.
            bool isVerifyR382 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "HomePhone",
                                            "{2ab923eb-9880-4b47-9965-ebf93ae15487}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R382
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR382,
                                            382,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]HomePhone[ Field.ID:]{2ab923eb-9880-4b47-9965-ebf93ae15487}[Field.Type:]Text.");

            // Verify JobTitle field's id and type.
            bool isVerifyR387 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "JobTitle",
                                            "{c4e0f350-52cc-4ede-904c-dd71a3d11f7d}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R387
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR387,
                                            387,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]JobTitle[ Field.ID:]{c4e0f350-52cc-4ede-904c-dd71a3d11f7d}[Field.Type:]Text.");

            // Verify LastNamePhonetic field's id and type.
            bool isVerifyR389 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "LastNamePhonetic",
                                            "{fdc8216d-dabf-441d-8ac0-f6c626fbdc24}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R389
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR389,
                                            389,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]LastNamePhonetic[ Field.ID:]{fdc8216d-dabf-441d-8ac0-f6c626fbdc24}[Field.Type:]Text.");

            // Verify Title field's id and type.
            bool isVerifyR419 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Title",
                                            "{fa564e0f-0c70-4ab9-b863-0177e6ddd247}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R419
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR419,
                                            419,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Title[ Field.ID:]{fa564e0f-0c70-4ab9-b863-0177e6ddd247}[Field.Type:]Text.");

            // Verify WebPage field's id and type.
            bool isVerifyR425 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "WebPage",
                                            "{a71affd2-dcc7-4529-81bc-2fe593154a5f}",
                                            "URL");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R425
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR425,
                                            425,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]WebPage[ Field.ID:]{a71affd2-dcc7-4529-81bc-2fe593154a5f}[Field.Type:]URL.");

            // Verify WorkAddress field's id and type.
            bool isVerifyR426 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "WorkAddress",
                                            "{fc2e188e-ba91-48c9-9dd3-16431afddd50}",
                                            "Note");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R426
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR426,
                                            426,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]WorkAddress[ Field.ID:]{fc2e188e-ba91-48c9-9dd3-16431afddd50}[Field.Type:]Note.");

            // Verify WorkCity field's id and type.
            bool isVerifyR427 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "WorkCity",
                                            "{6ca7bd7f-b490-402e-af1b-2813cf087b1e}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R427
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR427,
                                            427,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]WorkCity[ Field.ID:]{6ca7bd7f-b490-402e-af1b-2813cf087b1e}[Field.Type:]Text.");

            // Verify WorkCountry field's id and type.
            bool isVerifyR428 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "WorkCountry",
                                            "{3f3a5c85-9d5a-4663-b925-8b68a678ea3a}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R428
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR428,
                                            428,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]WorkCountry[ Field.ID:]{3f3a5c85-9d5a-4663-b925-8b68a678ea3a}[Field.Type:]Text.");

            // Verify WorkFax field's id and type.
            bool isVerifyR429 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "WorkFax",
                                            "{9d1cacc8-f452-4bc1-a751-050595ad96e1}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R429
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR429,
                                            429,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]WorkFax[ Field.ID:]{9d1cacc8-f452-4bc1-a751-050595ad96e1}[Field.Type:]Text.");

            // Verify WorkPhone field's id and type.
            bool isVerifyR431 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "WorkPhone",
                                            "{fd630629-c165-4513-b43c-fdb16b86a14d}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R431
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR431,
                                            431,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]WorkPhone[ Field.ID:]{fd630629-c165-4513-b43c-fdb16b86a14d}[Field.Type:]Text.");

            // Verify WorkState field's id and type.
            bool isVerifyR433 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "WorkState",
                                            "{ceac61d3-dda9-468b-b276-f4a6bb93f14f}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R433
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR433,
                                            433,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]WorkState[ Field.ID:]{ceac61d3-dda9-468b-b276-f4a6bb93f14f}[Field.Type:]Text.");

            // Verify WorkZip field's id and type.
            bool isVerifyR434 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "WorkZip",
                                            "{9a631556-3dac-49db-8d2f-fb033b0fdc24}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R434
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR434,
                                            434,
                                            @"[In Contact-Specific Schema][One of the contact properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]WorkZip[ Field.ID:]{9a631556-3dac-49db-8d2f-fb033b0fdc24}[Field.Type:]Text.");

            #endregion Verify fields
        }

        #endregion

        #region MSOUTSPS_S03_TC03_VerifyDiscussionList

        /// <summary>
        /// This test case is used to verify definition of Discussion list which contains contact list items.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S03_TC03_VerifyDiscussionList()
        {
            // Add a DiscussionBoard list.
            string listId = this.AddListToSUT(TemplateType.Discussion_Board);

            // Call GetList operation
            GetListResponseGetListResult getListResult = OutspsAdapter.GetList(listId);

            #region Verify fields

            // Verify Author field's id and type.
            bool isVerifyR684 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Author",
                                            "{1df5e554-ec7e-46a6-901d-d85a3881cb18}",
                                            "User");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R684
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR684,
                                            684,
                                            @"[In Discussion-Specific Schema][One of the discussion properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Author[ Field.ID:]{1df5e554-ec7e-46a6-901d-d85a3881cb18}[Field.Type:]User.");

            // Verify Body field's id and type.
            bool isVerifyR685 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Body",
                                            "{7662cd2c-f069-4dba-9e35-082cf976e170}",
                                            "Note");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R685
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR685,
                                            685,
                                            @"[In Discussion-Specific Schema][One of the discussion properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Body[ Field.ID:]{7662cd2c-f069-4dba-9e35-082cf976e170}[Field.Type:]Note.");

            // Verify DiscussionTitle field's id and type.
            bool isVerifyR686 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "DiscussionTitle",
                                            "{c5abfdc7-3435-4183-9207-3d1146895cf8}",
                                            "Computed");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R686
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR686,
                                            686,
                                            @"[In Discussion-Specific Schema][One of the discussion properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]DiscussionTitle[ Field.ID:]{c5abfdc7-3435-4183-9207-3d1146895cf8}[Field.Type:]Computed.");

            // Verify Editor field's id and type.
            bool isVerifyR687 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Editor",
                                            "{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}",
                                            "User");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R687
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR687,
                                            687,
                                            @"[In Discussion-Specific Schema][One of the discussion properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Editor[ Field.ID:]{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}[Field.Type:]User.");

            // Verify ThreadIndex field's id and type.
            bool isVerifyR688 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "ThreadIndex",
                                            "{cef73bf1-edf6-4dd9-9098-a07d83984700}",
                                            "ThreadIndex");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R688
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR688,
                                            688,
                                            @"[In Discussion-Specific Schema][One of the discussion properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]ThreadIndex[ Field.ID:]{cef73bf1-edf6-4dd9-9098-a07d83984700}[Field.Type:]ThreadIndex.");

            // Verify Title field's id and type.
            bool isVerifyR689 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Title",
                                            "{fa564e0f-0c70-4ab9-b863-0177e6ddd247}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R689
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR689,
                                            689,
                                            @"[In Discussion-Specific Schema][One of the discussion properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Title[ Field.ID:]{fa564e0f-0c70-4ab9-b863-0177e6ddd247}[Field.Type:]Text.");

            #endregion Verify fields
        }

        #endregion

        #region MSOUTSPS_S03_TC04_VerifyDocumentLibrary

        /// <summary>
        /// This test case is used to verify definition of DocumentLibrary which contains contact list items.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S03_TC04_VerifyDocumentLibrary()
        {
            // Add a Document Library.
            string listId = this.AddListToSUT(TemplateType.Document_Library);

            // Call GetList operation
            GetListResponseGetListResult getListResult = OutspsAdapter.GetList(listId);

            #region Verify fields

            // Verify Author field's id and type.
            bool isVerifyR690 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Author",
                                            "{1df5e554-ec7e-46a6-901d-d85a3881cb18}",
                                            "User");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R690
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR690,
                                            690,
                                            @"[In Document-Specific Schema]One of the document properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Author[ Field.ID:]{1df5e554-ec7e-46a6-901d-d85a3881cb18}[Field.Type:]User.");

            // Verify Editor field's id and type.
            bool isVerifyR691 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Editor",
                                            "{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}",
                                            "User");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R691
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR691,
                                            691,
                                            @"[In Document-Specific Schema][One of the document properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Editor[ Field.ID:]{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}[Field.Type:]User.");

            // Verify EncodedAbsUrl field's id and type.
            bool isVerifyR692 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "EncodedAbsUrl",
                                            "{7177cfc7-f399-4d4d-905d-37dd51bc90bf}",
                                            "Computed");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R692
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR692,
                                            692,
                                            @"[In Document-Specific Schema][One of the document properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]EncodedAbsUrl[ Field.ID:]{7177cfc7-f399-4d4d-905d-37dd51bc90bf}[Field.Type:]Computed.");

            // Verify fieldirRef field's id and type.
            bool isVerifyR693 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "FileDirRef",
                                            "{56605df6-8fa1-47e4-a04c-5b384d59609f}",
                                            "Lookup");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R693
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR693,
                                            693,
                                            @"[In Document-Specific Schema][One of the document properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]FileDirRef[ Field.ID:]{56605df6-8fa1-47e4-a04c-5b384d59609f}[Field.Type:]Lookup.");
                                              
            // Verify FileSizeDisplay field's id and type.
            bool isVerifyR694 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "FileSizeDisplay",
                                            "{78a07ba4-bda8-4357-9e0f-580d64487583}",
                                            "Computed");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R694
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR694,
                                            694,
                                            @"[In Document-Specific Schema][One of the document properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]FileSizeDisplay[ Field.ID:]{78a07ba4-bda8-4357-9e0f-580d64487583}[Field.Type:]Computed.");

            // Verify LinkCheckedOutTitle field's id and type.
            bool isVerifyR695 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "LinkCheckedOutTitle",
                                            "{e2a15dfd-6ab8-4aec-91ab-02f6b64045b0}",
                                            "Computed");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R695
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR695,
                                            695,
                                            @"[In Document-Specific Schema][One of the document properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]LinkCheckedOutTitle[ Field.ID:]{e2a15dfd-6ab8-4aec-91ab-02f6b64045b0}[Field.Type:]Computed.");

            // Verify LinkFilename field's id and type.
            bool isVerifyR696 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "LinkFilename",
                                            "{5cc6dc79-3710-4374-b433-61cb4a686c12}",
                                            "Computed");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R696
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR696,
                                            696,
                                            @"[In Document-Specific Schema][One of the document properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]LinkFilename[ Field.ID:]{5cc6dc79-3710-4374-b433-61cb4a686c12}[Field.Type:]Computed.");

            #endregion Verify fields
        }

        #endregion

        #region MSOUTSPS_S03_TC05_VerifyTasksList

        /// <summary>
        /// This test case is used to verify definition of Tasks list which contains contact list items.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S03_TC05_VerifyTasksList()
        {
            // Add a Tasks list.
            string listId = this.AddListToSUT(TemplateType.Tasks);

            // Call GetList operation
            GetListResponseGetListResult getListResult = OutspsAdapter.GetList(listId);

            #region Verify fields

            // Verify AssignedTo field's id and type.
            bool isVerifyR6983 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "AssignedTo",
                                            "{53101f38-dd2e-458c-b245-0c236cc13d1a}",
                                            "User");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R6983
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR6983,
                                            6983,
                                            @"[In Appendix B: Product Behavior] Implementation does set the default value of AssignedTo to User. (<37> Section 3.2.4.2.8:  For  SharePoint Server 2010 and above, [In Task-Specific Schema][One of the task properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]AssignedTo[ Field.ID:]{53101f38-dd2e-458c-b245-0c236cc13d1a}[Field.Type:]User.)");

            // Verify Body field's id and type.
            bool isVerifyR699 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Body",
                                            "{7662cd2c-f069-4dba-9e35-082cf976e170}",
                                            "Note");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R699
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR699,
                                            699,
                                            @"[In Task-Specific Schema][One of the task properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Body[ Field.ID:]{7662cd2c-f069-4dba-9e35-082cf976e170}[Field.Type:]Note.");

            // Verify DueDate field's id and type.
            bool isVerifyR702 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "DueDate",
                                            "{cd21b4c2-6841-4f9e-a23a-738a65f99889}",
                                            "DateTime");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R702
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR702,
                                            702,
                                            @"[In Task-Specific Schema][One of the task properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]DueDate[ Field.ID:]{cd21b4c2-6841-4f9e-a23a-738a65f99889}[Field.Type:]DateTime.");

            // Verify Editor field's id and type.
            bool isVerifyR703 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Editor",
                                            "{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}",
                                            "User");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R703
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR703,
                                            703,
                                            @"[In Task-Specific Schema][One of the task properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Editor[ Field.ID:]{d31655d1-1d5b-4511-95a1-7a09e9b75bf2}[Field.Type:]User.");

            // Verify PercentComplete field's id and type.
            bool isVerifyR705 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "PercentComplete",
                                            "{d2311440-1ed6-46ea-b46d-daa643dc3886}",
                                            "Number");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R705
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR705,
                                            705,
                                            @"[In Task-Specific Schema][One of the task properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]PercentComplete[ Field.ID:]{d2311440-1ed6-46ea-b46d-daa643dc3886}[Field.Type:]Number.");

            // Verify Priority field's id and type.
            bool isVerifyR706 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Priority",
                                            "{a8eb573e-9e11-481a-a8c9-1104a54b2fbd}",
                                            "Choice");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R706
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR706,
                                            706,
                                            @"[In Task-Specific Schema][One of the task properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Priority[ Field.ID:]{a8eb573e-9e11-481a-a8c9-1104a54b2fbd}[Field.Type:]Choice.");

            // Verify StartDate field's id and type.
            bool isVerifyR708 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "StartDate",
                                            "{64cd368d-2f95-4bfc-a1f9-8d4324ecb007}",
                                            "DateTime");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R708
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR708,
                                            708,
                                            @"[In Task-Specific Schema][One of the task properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]StartDate[ Field.ID:]{64cd368d-2f95-4bfc-a1f9-8d4324ecb007}[Field.Type:]DateTime.");

            // Verify Status field's id and type.
            bool isVerifyR709 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Status",
                                            "{c15b34c3-ce7d-490a-b133-3f4de8801b76}",
                                            "Choice");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R709
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR709,
                                            709,
                                            @"[In Task-Specific Schema][One of the task properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Status[ Field.ID:]{c15b34c3-ce7d-490a-b133-3f4de8801b76}[Field.Type:]Choice.");

            // Verify Title field's id and type.
            bool isVerifyR711 = this.VerifyFieldTypeAndId(
                                            getListResult,
                                            "Title",
                                            "{fa564e0f-0c70-4ab9-b863-0177e6ddd247}",
                                            "Text");

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R711
            this.Site.CaptureRequirementIfIsTrue(
                                            isVerifyR711,
                                            711,
                                            @"[In Task-Specific Schema][One of the task properties appears as the attributes of the element GetListResponse.GetListResult.List.Fields.Field. Field.Name:]Title[ Field.ID:]{fa564e0f-0c70-4ab9-b863-0177e6ddd247}[Field.Type:]Text.");

            #endregion Verify fields
        }

        #endregion

        #region MSOUTSPS_S03_TC06_VerifyCHOICESAndMAPPINGSElements_TasksList

        /// <summary>
        /// This test case is used to verify CHOICES and MAPPINGS element of Tasks list.
        /// </summary>
        [TestCategory("MSOUTSPS"), TestMethod()]
        public void MSOUTSPS_S03_TC06_VerifyCHOICESAndMAPPINGSElements_TasksList()
        {
            // Add a Tasks list.
            string listId = this.AddListToSUT(TemplateType.Tasks);

            // Call GetList operation and get the "Priority" field.
            GetListResponseGetListResult getListResult = OutspsAdapter.GetList(listId);
            FieldDefinition fieldOfPriority = Common.GetFieldItemByName(getListResult, "Priority", this.Site);
            XmlElement rawResponsOfGetList = SchemaValidation.LastRawResponseXml;

            // Verify schema, if pass the schema validation for CHOICES and MAPPINGS elements, capture R774, R775, R778
            bool isPassTheSchemaValidation = this.VerifyChoicesAndMappingsSchema(rawResponsOfGetList, "Priority");

            this.Site.CaptureRequirementIfIsTrue(
                                                isPassTheSchemaValidation,
                                                774,
                                                @"[In Task-Specific Schema]Each MAPPINGS element holds a number and a string.");

            this.Site.CaptureRequirementIfIsTrue(
                                               isPassTheSchemaValidation,
                                               775,
                                               @"[In Task-Specific Schema][The schema definition of CHOICES is:]<s:element name=""CHOICES"" >
   < s:sequence >
      < s:element name = ""CHOICE"" type = ""string"" minOccurs = ""0"" maxOccurs = ""unbounded"" />
   </ s:sequence >
</ s:element > ");

            this.Site.CaptureRequirementIfIsTrue(
                                               isPassTheSchemaValidation,
                                               778,
                                               @"[In Task-Specific Schema][The schema definiton of MAPPINGS is:]<s:complexType name=""MAPPINGS"">
   <s:sequence>
      <s:element name=""MAPPING"" type=""string"" minOccurs=""0"" maxOccurs=""unbounded"">
         <s:complexType>
            <s:simpleContent>
               <s:extension base=""string"">
                  <s:attribute name=""Value"" type=""integer"" use=""required"" />
               </s:extension>
            </s:simpleContent>
         </s:complexType>
      </s:element>
   </s:sequence>
</s:complexType>");

            // Verify relationship between the CHOICES and MAPPINGS elements, each CHOICE item should be found a mapped item in MAPPINGS array.
            bool isMatchRelationShip = this.VerifyChoicesAndMappingsRelationShip(fieldOfPriority);

            // If pass the relationship validation, capture R77301, R779
            if (Common.IsRequirementEnabled(77301, this.Site))
            {
                // Verify MS-OUTSPS requirement: MS-OUTSPS_R77301
                this.Site.CaptureRequirementIfIsTrue(
                                                  isMatchRelationShip,
                                                  77301,
                                                  @"[In Task-Specific Schema]Implementation does provide a MAPPINGS element.(Windows® SharePoint® Services 3.0 and above products follow this behavior.)");
            }

            // Verify MS-OUTSPS requirement: MS-OUTSPS_R779
            this.Site.CaptureRequirementIfIsTrue(
                                               isMatchRelationShip,
                                               779,
                                               @"[In Task-Specific Schema]MAPPINGS.MAPPING and MAPPINGS.Value: If a string in a CHOICE element exactly matches a string in a MAPPING element, then the MAPPING.Value attribute associated with it can be used to represent the string.");

            // Call GetList operation and get the "Status" field.
            FieldDefinition fieldOfStatus = Common.GetFieldItemByName(getListResult, "Status", this.Site);

            isPassTheSchemaValidation = this.VerifyChoicesAndMappingsSchema(rawResponsOfGetList, "Status");
            this.Site.Assert.IsTrue(isPassTheSchemaValidation, "The MAPPINGS and CHOICES elements should match the schema definition.");
            
            isMatchRelationShip = this.VerifyChoicesAndMappingsRelationShip(fieldOfStatus);
            this.Site.Assert.IsTrue(isMatchRelationShip, "Each CHOICE item should be found a mapped item in MAPPINGS array");

            // If pass the relationship validation, and pass schema validation for CHOICES and MAPPINGS elements, capture R784
            this.Site.CaptureRequirement(
                                    784,
                                    @"[In Task-Specific Schema]This property[Status] works exactly the same way as the Priority task property also specified in this section (see section 3.2.4.2.8).");
        }

        #endregion

        #endregion Test cases

        #region private methods

        /// <summary>
        /// A method used to verify field's id or field's type is equal to expected value.
        /// </summary>
        /// <param name="getListResponse">A parameter represents the response of GetList operation which contains the field definitions.</param>
        /// <param name="fieldName">A parameter represents the name of a field definition which is used to get the definition from the  response of GetList operation.</param>
        /// <param name="expectedFieldId">A parameter represents the expected id of field definition.</param>
        /// <param name="expectedFieldType">A parameter represents the expected type of field definition.</param>
        /// <returns>Return true indicating the verification pass.</returns>
        private bool VerifyFieldTypeAndId(GetListResponseGetListResult getListResponse, string fieldName, string expectedFieldId, string expectedFieldType)
        {
            if (string.IsNullOrEmpty(fieldName))
            {
                throw new ArgumentNullException("fieldName");
            }

            // Get the field definition from response.
            FieldDefinition fieldDefinition = Common.GetFieldItemByName(getListResponse, fieldName, this.Site);

            // Verify field's type and field's value
            bool isEqualToExpectedFieldType = false;
            bool isEqualToExpectedFieldId = false;

            // Ignore the field type verification if not specified expected value.
            bool isIgnoreFieldType = string.IsNullOrEmpty(expectedFieldType);
            if (isIgnoreFieldType)
            {
                isEqualToExpectedFieldType = true;
            }
            else
            {
                isEqualToExpectedFieldType = Common.VerifyFieldType(fieldDefinition, expectedFieldType, this.Site);
            }

            // Ignore the field Id verification if not specified expected value.
            bool isIgnoreFieldId = string.IsNullOrEmpty(expectedFieldId);
            if (isIgnoreFieldId)
            {
                isEqualToExpectedFieldId = true;
            }
            else
            {
                isEqualToExpectedFieldId = Common.VerifyFieldId(fieldDefinition, expectedFieldId, this.Site);
            }

            bool isPassVerification = isEqualToExpectedFieldType && isEqualToExpectedFieldId;

            // Add logs
            string verificationMsgOfType = string.Format(
                                "Field type verification perform:[{0}]; actual type value:[{1}], expected value[{2}]",
                                isIgnoreFieldType ? "No" : "Yes",
                                string.IsNullOrEmpty(fieldDefinition.Type) ? @"N/A" : fieldDefinition.Type,
                                isIgnoreFieldType ? @"N/A" : expectedFieldType);

             string verificationMsgOfId = string.Format(
                                "Field id verification perform:[{0}]; actual id value[{1}], expected value[{2}]",
                                isIgnoreFieldId ? "No" : "Yes",
                                string.IsNullOrEmpty(fieldDefinition.ID) ? @"N/A" : fieldDefinition.ID,
                                isIgnoreFieldId ? @"N/A" : expectedFieldId);

            this.Site.Log.Add(
                    LogEntryKind.Debug,
                    "Verification detail:\r\n{0}\r\n{1}",
                    verificationMsgOfType,
                    verificationMsgOfId);

            return isPassVerification;
        }

        /// <summary>
        /// A method used to verify the relationship between CHOICES and MAPPINGS elements. If the CHOICES element present, the protocol SUT should provide a MAPPINGS element. Each CHOICE element should found a mapped MAPPING item.
        /// </summary>
        /// <param name="fieldDefinition">A parameter represents the FieldDefinition instance which contains the CHOICES and MAPPINGS elements.</param>
        /// <returns>Return 'true' indicating the relationship validation pass.</returns>
        private bool VerifyChoicesAndMappingsRelationShip(FieldDefinition fieldDefinition)
        {
           if (null == fieldDefinition.MAPPINGS)
           {
               throw new ArgumentException("The fieldDfinition should contain valid MAPPINGDEFINITION[] type instance.", string.Empty);
           }

           if (null == fieldDefinition.CHOICES)
           {
               throw new ArgumentException("The fieldDfinition should contain valid MAPPINGDEFINITION[] type instance.");
           }

           if (0 == fieldDefinition.MAPPINGS.Length || 0 == fieldDefinition.CHOICES.Length)
           {
               // If the MAPPINGS and CHOICES element does not contain any item, it is match the TD.
               return true;
           }

           if (fieldDefinition.MAPPINGS.Length != fieldDefinition.CHOICES.Length)
           {
               this.Site.Assert.Fail("The MAPPING items' number should equal to the CHOICE items' number.");
           }

           bool isMappingItemsMatchChoicesItems = false;
           foreach (CHOICEDEFINITION choiceItem in fieldDefinition.CHOICES)
           {
               string choiceItemValue = choiceItem.Text[0];

               // The choice item Value should be found in MAPPINGS array.
               isMappingItemsMatchChoicesItems = fieldDefinition.MAPPINGS.Any(Founder => Founder.Value1.Equals(choiceItemValue, StringComparison.OrdinalIgnoreCase));
               if (!isMappingItemsMatchChoicesItems)
               {
                   break;
               }
           }

           return isMappingItemsMatchChoicesItems;
        }

        /// <summary>
        /// A method used to verify the schema definition for CHOICES and MAPPINGS elements. If there are any errors, method will throw a schema validation exception.
        /// </summary>
        /// <param name="rawResponseOfGetList">A parameter represents the raw response of GetList operation. This value could be get from "LastRawResponseXml" property of protocol adapter.</param>
        /// <param name="expectedFieldName">A parameter represents the name of the field definition which the method will look up the CHOICES and Mappings elements xml string.</param>
        /// <returns>Return 'true' indicating the schema validation pass.</returns>
        private bool VerifyChoicesAndMappingsSchema(XmlElement rawResponseOfGetList, string expectedFieldName)
        {
            // Extract the Choices and Mapping xml string from response of GetList operation.
            List<string> choicesAndMappingXmlstrings = this.GetChoicesAndMappingsXmlString(rawResponseOfGetList, expectedFieldName);

            // Verify the schema definition of Choices and Mapping elements.
            foreach (string elementXmlString in choicesAndMappingXmlstrings)
            {
                SchemaValidation.ValidateXml(this.Site, elementXmlString);
                if (SchemaValidation.ValidationResult != ValidationResult.Success)
                {
                    string validationErrorMessage = SchemaValidation.GenerateValidationResult();
                    throw new XmlSchemaValidationException(validationErrorMessage);
                }
            }

            // If there are no any schema validation exception thrown, return true.
            return true;
        }

        /// <summary>
        /// A method used to get the CHOICES and Mappings elements xml string from the raw response of GetList operation.
        /// </summary>
        /// <param name="rawResponseOfGetList">A parameter represents the raw response of GetList operation. This value is the xml string.</param>
        /// <param name="expectedFieldName">A parameter represents the name of the field definition which the method will look up the CHOICES and Mappings elements xml string.</param>
        /// <returns>A return value represents the xml string collection contains both CHOICES and MAPPINGS element. The first item of the xml string collection is the CHOICES element xml string, the second item is the MAPPINGS element xml string.</returns>
        private List<string> GetChoicesAndMappingsXmlString(XmlElement rawResponseOfGetList, string expectedFieldName)
        {
            if (string.IsNullOrEmpty(expectedFieldName))
            {
                throw new ArgumentNullException("expectedFieldName");
            }

            if (null == rawResponseOfGetList)
            {
                throw new ArgumentNullException("rawResponseOfGetList");
            }

            XmlNodeList getListResultElement = rawResponseOfGetList.GetElementsByTagName("GetListResult");
            if (0 == getListResultElement.Count)
            {
                throw new ArgumentException("The raw response should be GetList operation, it must contain GetListResult element.");
            }

            XmlNodeList fieldDefinitionitems = rawResponseOfGetList.GetElementsByTagName("Field");
            var priorityItems = from XmlElement fieldDefinitionItem in fieldDefinitionitems
                                where "Priority".Equals(fieldDefinitionItem.Attributes["Name"].Value, StringComparison.OrdinalIgnoreCase)
                                select fieldDefinitionItem;

            this.Site.Assert.AreEqual<int>(
                                        1,
                                        priorityItems.Count(),
                                        "The response of GetList operation for task list should contain [{0}] field definition.",
                                        expectedFieldName);

            XmlElement expectedFieldElement = priorityItems.ElementAt<XmlElement>(0);
            List<string> elementXmlStrings = new List<string>();

            XmlNodeList choicesElements = expectedFieldElement.GetElementsByTagName("CHOICES");
            this.Site.Assert.AreEqual<int>(
                                    1,
                                    choicesElements.Count,
                                    "The [{0}] field should contain 'CHOICES' element.",
                                    expectedFieldName);
            elementXmlStrings.Add(choicesElements[0].OuterXml);

            XmlNodeList mappingsElements = expectedFieldElement.GetElementsByTagName("MAPPINGS");
            this.Site.Assert.AreEqual<int>(
                                    1,
                                    mappingsElements.Count,
                                    "The [{0}] field should contain 'MAPPINGS' element.",
                                    expectedFieldName);
            elementXmlStrings.Add(mappingsElements[0].OuterXml);

            return elementXmlStrings;
        }

        #endregion private methods
    }
}
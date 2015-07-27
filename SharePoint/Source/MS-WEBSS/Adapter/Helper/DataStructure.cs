//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.MS_WEBSS
{
    #region User Authentication

    /// <summary>
    /// This enumeration indicate the option of authentication information client used.
    /// </summary>
    public enum UserAuthentication
    {
        /// <summary>
        /// Specify Adapter use a authenticated account.
        /// </summary>
        Authenticated,

        /// <summary>
        /// Specify Adapter use a unauthenticated account.
        /// </summary>
        Unauthenticated
    }

    #endregion

    /// <summary>
    /// Const chars.
    /// </summary>
    public enum CONST_CHARS : int
    {
        /// <summary>
        /// const char:Period
        /// </summary>
        Period = (int)'.',

        /// <summary>
        /// const char:Tab
        /// </summary>
        Tab = (int)'\t',

        /// <summary>
        /// const char:Backslash
        /// </summary>
        Backslash = (int)'\\',

        /// <summary>
        /// const char:Slash
        /// </summary>
        Slash = (int)'/',

        /// <summary>
        /// const char:Colon
        /// </summary>
        Colon = (int)':',

        /// <summary>
        /// const char:Asterisk
        /// </summary>
        Asterisk = (int)'*',

        /// <summary>
        /// const char:QuestionMark
        /// </summary>
        QuestionMark = (int)'?',

        /// <summary>
        /// const char:QuotationMark
        /// </summary>
        QuotationMark = (int)'"',

        /// <summary>
        /// const char:NumberSign
        /// </summary>
        NumberSign = (int)'#',

        /// <summary>
        /// const char:PercentSign
        /// </summary>
        PercentSign = (int)'%',

        /// <summary>
        /// const char:LessThanSign
        /// </summary>
        LessThanSign = (int)'<',

        /// <summary>
        /// const char:GreaterThanSign
        /// </summary>
        GreaterThanSign = (int)'>',

        /// <summary>
        /// const char:OpeningCurlyBraces
        /// </summary>
        OpeningCurlyBraces = (int)'{',

        /// <summary>
        /// const char:ClosingCurlyBraces
        /// </summary>
        ClosingCurlyBraces = (int)'}',

        /// <summary>
        /// const character VerticalBar
        /// </summary>
        VerticalBar = (int)'|',

        /// <summary>
        /// const character Tilde
        /// </summary>
        Tilde = (int)'~',

        /// <summary>
        /// const character Ampersand
        /// </summary>
        Ampersand = (int)'&',

        /// <summary>
        /// const character Comma
        /// </summary>
        Comma = (int)',',

        /// <summary>
        /// const character Blank
        /// </summary>
        Blank = (int)' '
    }

    /// <summary>
    /// CONST_MethodID UNITS
    /// </summary>
    public enum CONST_MethodIDUNITS : int
    {
        /// <summary>
        /// Specify Method ID is 1.
        /// </summary>
        One = 1,

        /// <summary>
        /// Specify Method ID is 2.
        /// </summary>
        Two = 2,

        /// <summary>
        /// Specify Method ID is 3
        /// </summary>
        Three = 3,

        /// <summary>
        /// Specify Method ID is 4
        /// </summary>
        Four = 4,

        /// <summary>
        /// Specify Method ID is 5
        /// </summary>
        Five = 5,

        /// <summary>
        /// Specify Method ID is 6
        /// </summary>
        Six = 6
    }

    /// <summary>
    /// LCID values
    /// </summary>
    public enum LCID_Values : int
    {
        /// <summary>
        /// Indicates Afrikaans.
        /// </summary>
        Afrikaans = 1078,

        /// <summary>
        /// Indicates Albanian.
        /// </summary>
        Albanian = 1052,

        /// <summary>
        /// Indicates Arabic_United_Arab_Emirates.
        /// </summary>
        Arabic_United_Arab_Emirates = 14337,

        /// <summary>
        /// Indicates Arabic_Bahrain.
        /// </summary>
        Arabic_Bahrain = 15361,

        /// <summary>
        /// Indicates Arabic_Algeria.
        /// </summary>
        Arabic_Algeria = 5121,

        /// <summary>
        /// Indicates Arabic_Egypt.
        /// </summary>
        Arabic_Egypt = 3073,

        /// <summary>
        /// Indicates Arabic_Iraq.
        /// </summary>
        Arabic_Iraq = 2049,

        /// <summary>
        /// Indicates Arabic_Jordan.
        /// </summary>
        Arabic_Jordan = 11265,

        /// <summary>
        /// Indicates Arabic_Kuwait.
        /// </summary>
        Arabic_Kuwait = 13313,

        /// <summary>
        /// Indicates Arabic_Lebanon.
        /// </summary>
        Arabic_Lebanon = 12289,

        /// <summary>
        /// Indicates Arabic_Libya.
        /// </summary>
        Arabic_Libya = 4097,

        /// <summary>
        /// Indicates Arabic_Morocco.
        /// </summary>
        Arabic_Morocco = 6145,

        /// <summary>
        /// Indicates Arabic_Oman.
        /// </summary>
        Arabic_Oman = 8193,

        /// <summary>
        /// Indicates Arabic_Qatar.
        /// </summary>
        Arabic_Qatar = 16385,

        /// <summary>
        /// Indicates Arabic_Saudi_Arabia.
        /// </summary>
        Arabic_Saudi_Arabia = 1025,

        /// <summary>
        /// Indicates Arabic_Syria.
        /// </summary>
        Arabic_Syria = 10241,

        /// <summary>
        /// Indicates Arabic_Tunisia.
        /// </summary>
        Arabic_Tunisia = 7169,

        /// <summary>
        /// Indicates Arabic_Yemen.
        /// </summary>
        Arabic_Yemen = 9217,

        /// <summary>
        /// Indicates Armenian.
        /// </summary>
        Armenian = 1067,

        /// <summary>
        /// Indicates Azeri_Latin.
        /// </summary>
        Azeri_Latin = 1068,

        /// <summary>
        /// Indicates Azeri_Cyrillic.
        /// </summary>
        Azeri_Cyrillic = 2092,

        /// <summary>
        /// Indicates Basque (Basque).
        /// </summary>
        Basque = 1069,

        /// <summary>
        /// Indicates Belarusian.
        /// </summary>
        Belarusian = 1059,

        /// <summary>
        /// Indicates Bulgarian.
        /// </summary>
        Bulgarian = 1026,

        /// <summary>
        /// Indicates Catalan.
        /// </summary>
        Catalan = 1027,

        /// <summary>
        /// Indicates Chinese_China.
        /// </summary>
        Chinese_China = 2052,

        /// <summary>
        /// Indicates Chinese_Hong_Kong_SAR.
        /// </summary>
        Chinese_Hong_Kong_SAR = 3076,

        /// <summary>
        /// Indicates Chinese_Macau_SAR.
        /// </summary>
        Chinese_Macau_SAR = 5124,

        /// <summary>
        /// Indicates Chinese_Singapore.
        /// </summary>
        Chinese_Singapore = 4100,

        /// <summary>
        /// Indicates Chinese (Traditional).
        /// </summary>
        Chinese_Taiwan = 1028,

        /// <summary>
        /// Indicates Croatian.
        /// </summary>
        Croatian = 1050,

        /// <summary>
        /// Indicates Czech.
        /// </summary>
        Czech = 1029,

        /// <summary>
        /// Indicates Danish.
        /// </summary>
        Danish = 1030,

        /// <summary>
        /// Indicates Dutch_Netherlands.
        /// </summary>
        Dutch_Netherlands = 1043,

        /// <summary>
        /// Indicates Dutch_Belgium.
        /// </summary>
        Dutch_Belgium = 2067,

        /// <summary>
        /// Indicates English_Australia.
        /// </summary>
        English_Australia = 3081,

        /// <summary>
        /// Indicates English_Belize.
        /// </summary>
        English_Belize = 10249,

        /// <summary>
        /// Indicates English_Canada.
        /// </summary>
        English_Canada = 4105,

        /// <summary>
        /// Indicates English_Caribbean.
        /// </summary>
        English_Caribbean = 9225,

        /// <summary>
        /// Indicates English_Ireland.
        /// </summary>
        English_Ireland = 6153,

        /// <summary>
        /// Indicates English_Jamaica.
        /// </summary>
        English_Jamaica = 8201,

        /// <summary>
        /// Indicates English_New_Zealand.
        /// </summary>
        English_New_Zealand = 5129,

        /// <summary>
        /// Indicates English_Philippines.
        /// </summary>
        English_Philippines = 13321,

        /// <summary>
        /// Indicates English_Southern_Africa.
        /// </summary>
        English_Southern_Africa = 7177,

        /// <summary>
        /// Indicates English_Trinidad.
        /// </summary>
        English_Trinidad = 11273,

        /// <summary>
        /// Indicates English_Great_Britain.
        /// </summary>
        English_Great_Britain = 2057,

        /// <summary>
        /// Indicates English_United_States.
        /// </summary>
        English_United_States = 1033,

        /// <summary>
        /// Indicates Estonian.
        /// </summary>
        Estonian = 1061,

        /// <summary>
        /// Indicates Farsi.
        /// </summary>
        Farsi = 1065,

        /// <summary>
        /// Indicates Finnish.
        /// </summary>
        Finnish = 1035,

        /// <summary>
        /// Indicates Faroese.
        /// </summary>
        Faroese = 1080,

        /// <summary>
        /// Indicates French_France.
        /// </summary>
        French_France = 1036,

        /// <summary>
        /// Indicates French_Belgium.
        /// </summary>
        French_Belgium = 2060,

        /// <summary>
        /// Indicates French_Canada.
        /// </summary>
        French_Canada = 3084,

        /// <summary>
        /// Indicates French_Luxembourg.
        /// </summary>
        French_Luxembourg = 5132,

        /// <summary>
        /// Indicates French_Switzerland.
        /// </summary>
        French_Switzerland = 4108,

        /// <summary>
        /// Indicates Gaelic_Ireland.
        /// </summary>
        Gaelic_Ireland = 2108,

        /// <summary>
        /// Indicates Gaelic_Scotland.
        /// </summary>
        Gaelic_Scotland = 1084,

        /// <summary>
        /// Indicates German_Germany.
        /// </summary>
        German_Germany = 1031,

        /// <summary>
        /// Indicates German_Austria.
        /// </summary>
        German_Austria = 3079,

        /// <summary>
        /// Indicates German_Liechtenstein.
        /// </summary>
        German_Liechtenstein = 5127,

        /// <summary>
        /// Indicates German_Luxembourg.
        /// </summary>
        German_Luxembourg = 4103,

        /// <summary>
        /// Indicates German_Switzerland.
        /// </summary>
        German_Switzerland = 2055,

        /// <summary>
        /// Indicates Greek.
        /// </summary>
        Greek = 1032,

        /// <summary>
        /// Indicates Hebrew.
        /// </summary>
        Hebrew = 1037,

        /// <summary>
        /// Indicates Hindi.
        /// </summary>
        Hindi = 1081,

        /// <summary>
        /// Indicates Hungarian.
        /// </summary>
        Hungarian = 1038,

        /// <summary>
        /// Indicates Icelandic.
        /// </summary>
        Icelandic = 1039,

        /// <summary>
        /// Indicates Indonesian.
        /// </summary>
        Indonesian = 1057,

        /// <summary>
        /// Indicates Italian_Italy.
        /// </summary>
        Italian_Italy = 1040,

        /// <summary>
        /// Indicates Italian_Switzerland.
        /// </summary>
        Italian_Switzerland = 2064,

        /// <summary>
        /// Indicates Japanese.
        /// </summary>
        Japanese = 1041,

        /// <summary>
        /// Indicates Korean.
        /// </summary>
        Korean = 1042,

        /// <summary>
        /// Indicates Latvian.
        /// </summary>
        Latvian = 1062,

        /// <summary>
        /// Indicates Lithuanian.
        /// </summary>
        Lithuanian = 1063,

        /// <summary>
        /// Indicates FYRO_Macedonia.
        /// </summary>
        FYRO_Macedonia = 1071,

        /// <summary>
        /// Indicates Malay_Malaysia.
        /// </summary>
        Malay_Malaysia = 1086,

        /// <summary>
        /// Indicates Malay_Brunei.
        /// </summary>
        Malay_Brunei = 2110,

        /// <summary>
        /// Indicates Maltese.
        /// </summary>
        Maltese = 1082,

        /// <summary>
        /// Indicates Marathi.
        /// </summary>
        Marathi = 1102,

        /// <summary>
        /// Indicates Norwegian_Bokml.
        /// </summary>
        Norwegian_Bokml = 1044,

        /// <summary>
        /// Indicates Norwegian_Nynorsk.
        /// </summary>
        Norwegian_Nynorsk = 2068,

        /// <summary>
        /// Indicates Polish.
        /// </summary>
        Polish = 1045,

        /// <summary>
        /// Indicates Portuguese_Portugal.
        /// </summary>
        Portuguese_Portugal = 2070,

        /// <summary>
        /// Indicates Portuguese_Brazil.
        /// </summary>
        Portuguese_Brazil = 1046,

        /// <summary>
        /// Indicates Raeto_Romance.
        /// </summary>
        Raeto_Romance = 1047,

        /// <summary>
        /// Indicates Romanian_Romania.
        /// </summary>
        Romanian_Romania = 1048,

        /// <summary>
        /// Indicates Romanian_Republic_of_Moldova.
        /// </summary>
        Romanian_Republic_of_Moldova = 2072,

        /// <summary>
        /// Indicates Russian.
        /// </summary>
        Russian = 1049,

        /// <summary>
        /// Indicates Russian_Republic_of_Moldova.
        /// </summary>
        Russian_Republic_of_Moldova = 2073,

        /// <summary>
        /// Indicates Sanskrit.
        /// </summary>
        Sanskrit = 1103,

        /// <summary>
        /// Indicates Serbian_Cyrillic.
        /// </summary>
        Serbian_Cyrillic = 3098,

        /// <summary>
        /// Indicates Serbian_Latin.
        /// </summary>
        Serbian_Latin = 2074,

        /// <summary>
        /// Indicates Setswana.
        /// </summary>
        Setswana = 1074,

        /// <summary>
        /// Indicates Slovenian.
        /// </summary>
        Slovenian = 1060,

        /// <summary>
        /// Indicates Slovak.
        /// </summary>
        Slovak = 1051,

        /// <summary>
        /// Indicates Sorbian.
        /// </summary>
        Sorbian = 1070,

        /// <summary>
        /// Indicates Spanish_Spain_Traditional.
        /// </summary>
        Spanish_Spain_Traditional = 1034,

        /// <summary>
        /// Indicates Spanish_Argentina.
        /// </summary>
        Spanish_Argentina = 11274,

        /// <summary>
        /// Indicates Spanish_Bolivia.
        /// </summary>
        Spanish_Bolivia = 16394,

        /// <summary>
        /// Indicates Spanish_Chile.
        /// </summary>
        Spanish_Chile = 13322,

        /// <summary>
        /// Indicates Spanish_Colombia.
        /// </summary>
        Spanish_Colombia = 9226,

        /// <summary>
        /// Indicates Spanish_Costa_Rica.
        /// </summary>
        Spanish_Costa_Rica = 5130,

        /// <summary>
        /// Indicates Spanish_Dominican_Republic.
        /// </summary>
        Spanish_Dominican_Republic = 7178,

        /// <summary>
        /// Indicates Spanish_Ecuador.
        /// </summary>
        Spanish_Ecuador = 12298,

        /// <summary>
        /// Indicates Spanish_Guatemala.
        /// </summary>
        Spanish_Guatemala = 4106,

        /// <summary>
        /// Indicates Spanish_Honduras.
        /// </summary>
        Spanish_Honduras = 18442,

        /// <summary>
        /// Indicates Spanish_Mexico.
        /// </summary>
        Spanish_Mexico = 2058,

        /// <summary>
        /// Indicates Spanish_Nicaragua.
        /// </summary>
        Spanish_Nicaragua = 19466,

        /// <summary>
        /// Indicates Spanish_Panama.
        /// </summary>
        Spanish_Panama = 6154,

        /// <summary>
        /// Indicates Spanish_Peru.
        /// </summary>
        Spanish_Peru = 10250,

        /// <summary>
        /// Indicates Spanish_Puerto_Rico.
        /// </summary>
        Spanish_Puerto_Rico = 20490,

        /// <summary>
        /// Indicates Spanish_Paraguay.
        /// </summary>
        Spanish_Paraguay = 15370,

        /// <summary>
        /// Indicates Spanish_El_Salvador.
        /// </summary>
        Spanish_El_Salvador = 17418,

        /// <summary>
        /// Indicates Spanish_Uruguay.
        /// </summary>
        Spanish_Uruguay = 14346,

        /// <summary>
        /// Indicates Spanish_Venezuela.
        /// </summary>
        Spanish_Venezuela = 8202,

        /// <summary>
        /// Indicates Southern_Sotho.
        /// </summary>
        Southern_Sotho = 1072,

        /// <summary>
        /// Indicates Swahili.
        /// </summary>
        Swahili = 1089,

        /// <summary>
        /// Indicates Swedish_Sweden.
        /// </summary>
        Swedish_Sweden = 1053,

        /// <summary>
        /// Indicates Swedish_Finland.
        /// </summary>
        Swedish_Finland = 2077,

        /// <summary>
        /// Indicates Tamil.
        /// </summary>
        Tamil = 1097,

        /// <summary>
        /// Indicates Tatar.
        /// </summary>
        Tatar = 1092,

        /// <summary>
        /// Indicates Thai.
        /// </summary>
        Thai = 1054,

        /// <summary>
        /// Indicates Turkish.
        /// </summary>
        Turkish = 1055,

        /// <summary>
        /// Indicates Tsonga.
        /// </summary>
        Tsonga = 1073,

        /// <summary>
        /// Indicates Ukrainian.
        /// </summary>
        Ukrainian = 1058,

        /// <summary>
        /// Indicates Urdu.
        /// </summary>
        Urdu = 1056,

        /// <summary>
        /// Indicates Uzbek_Cyrillic.
        /// </summary>
        Uzbek_Cyrillic = 2115,

        /// <summary>
        /// Indicates Uzbek_Latin.
        /// </summary>
        Uzbek_Latin = 1091,

        /// <summary>
        /// Indicates Vietnamese.
        /// </summary>
        Vietnamese = 1066,

        /// <summary>
        /// Indicates Xhosa.
        /// </summary>
        Xhosa = 1076,

        /// <summary>
        /// Indicates Yiddish.
        /// </summary>
        Yiddish = 1085,

        /// <summary>
        /// Indicates Zulu.
        /// </summary>
        Zulu = 1077
    }

    #region UpdateColumns

    /// <summary>
    /// This enumeration indicates the option of newFields parameter of UpdateColumns operation. 
    /// newFields is an XML element that represents the collection of columns to be added to the context site and all child sites within its hierarchy.
    /// </summary>
    public enum UpdateColumnsNewFieldsOption
    {
        /// <summary>
        /// Specify a valid field definition to be added
        /// </summary>
        Valid,

        /// <summary>
        /// Specify an invalid field definition to be added
        /// </summary>
        Invalid,

        /// <summary>
        /// Specify that request has multiple Method elements, without a Fields element defined as the root element.
        /// </summary>
        MultipleMethodNoFields,

        /// <summary>
        /// Specify an already existing column is wanted to be added to the site.
        /// </summary>
        ExistingColumn,

        /// <summary>
        /// Specify that neither Name nor DisplayName XML attributes are passed in a FieldDefinition element.
        /// </summary>
        NoNameOrDisplayName,

        /// <summary>
        /// Specify an invalid FieldDefinition element passed
        /// </summary>
        InvalidFieldDefinition,
    }

    /// <summary>
    /// This enumeration indicates the option of updateFields parameter of UpdateColumns operation. 
    /// updateFields is an XML element that represents the collection of columns to be updated on the context site and all child sites within its hierarchy
    /// </summary>
    public enum UpdateColumnsUpdateFieldsOption
    {
        /// <summary>
        /// Specify a valid field definition to be updated
        /// </summary>
        Valid,

        /// <summary>
        /// Specify an invalid field definition to be updated
        /// </summary>
        Invalid,

        /// <summary>
        /// Specify that request has multiple Method elements, without a Fields element defined as the root element.
        /// </summary>
        MultipleMethodNoFields,

        /// <summary>
        /// Specify that an invalid GUID is passed in as the ID XML attribute.
        /// </summary>
        InvalidGUID,

        /// <summary>
        /// Specify an invalid FieldDefinition element passed
        /// </summary>
        InvalidFieldDefinition,

        /// <summary>
        /// Specify the given Name XML attribute does not match to any of the already existing columns
        /// </summary>
        NoMatchingName,
    }

    /// <summary>
    /// This enumeration indicates the option of deleteFields parameter of UpdateColumns operation. 
    /// deleteFields is an XML element that represents the collection of columns to be deleted from the context site and all child sites within its hierarchy
    /// </summary>
    public enum UpdateColumnsDeleteFieldsOption
    {
        /// <summary>
        /// Specify a valid field definition to be deleted
        /// </summary>
        Valid,

        /// <summary>
        /// Specify an invalid field definition to be deleted
        /// </summary>
        Invalid,

        /// <summary>
        /// Specify that request has multiple Method elements, without a Fields element defined as the root element.
        /// </summary>
        MultipleMethodNoFields,

        /// <summary>
        /// Specify that a non-existing column is wanted to be deleted from the site.
        /// </summary>
        NonexistingColumn,

        /// <summary>
        /// Specify that an invalid GUID is passed in as the ID XML attribute.
        /// </summary>
        InvalidGUID,

        /// <summary>
        /// Specify an invalid FieldDefinition element passed
        /// </summary>
        InvalidFieldDefinition,

        /// <summary>
        /// Specify the given Name XML attribute does not match to any of the already existing columns
        /// </summary>
        NoMatchingName
    }

    #endregion

    /// <summary>
    /// The element names
    /// </summary>
    public struct ConstString
    {
        /// <summary>
        /// Indicates the name of Detail element.
        /// </summary>
        public const string Detail = "detail";

        /// <summary>
        /// Indicates ErrorCode.
        /// </summary>
        public const string ErrorCode = "errorcode";

        /// <summary>
        /// Indicates ErrorString.
        /// </summary>
        public const string ErrorString = "errorstring";

        /// <summary>
        /// Server relative URL.
        /// </summary>
        public const string DefaultServerUrl = "..";
    }

    /// <summary>
    /// The error code in the SOAP fault.
    /// </summary>
    public struct SoapErrorCode
    {
        /// <summary>
        /// Error code 0x82000001.
        /// </summary>
        public const string ErrorCode0x82000001 = "0x82000001";

        /// <summary>
        /// Error code 0x80131600.
        /// </summary>
        public const string ErrorCode0x80131600 = "0x80131600";

        /// <summary>
        /// Error code 0x80070002.
        /// </summary>
        public const string ErrorCode0x80070002 = "0x80070002";

        /// <summary>
        /// Error code 0x81020073.
        /// </summary>
        public const string ErrorCode0x81020073 = "0x81020073";

        /// <summary>
        /// Error code 0x00000000.
        /// </summary>
        public const string ErrorCode0x00000000 = "0x00000000";

        /// <summary>
        /// Error code 0x80004005.
        /// </summary>
        public const string ErrorCode0x80004005 = "0x80004005";

        /// <summary>
        /// Error code 0x82000007.
        /// </summary>
        public const string ErrorCode0x82000007 = "0x82000007";
    }

    /// <summary>
    /// The features' position
    /// </summary>
    public struct FeaturesPosition
    {
        /// <summary>
        /// The features on the site.
        /// </summary>
        public const string SiteFeatures = "site_features";

        /// <summary>
        /// The features on the parent site collection.
        /// </summary>
        public const string SiteCollectionFeatures = "site_collection_features";
    }
}
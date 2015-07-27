//-----------------------------------------------------------------------
// Copyright (c) 2014 Microsoft Corporation. All rights reserved.
// Use of this sample source code is subject to the terms of the Microsoft license 
// agreement under which you licensed this sample source code and is provided AS-IS.
// If you did not accept the terms of the license agreement, you are not authorized 
// to use this sample source code. For the terms of the license, please see the 
// license agreement between you and Microsoft.
//-----------------------------------------------------------------------

namespace Microsoft.Protocols.TestSuites.Common
{
    using System.Text.RegularExpressions;

    /// <summary>
    /// This class implements rfc 822 compliant email validator routines.
    /// </summary>
    public static class RFC822AddressParser
    {
        /// <summary>
        /// The constant string for the Escape character
        /// </summary>
        private const string Escape = @"\\";

        /// <summary>
        /// The constant string for the Period
        /// </summary>
        private const string Period = @"\.";

        /// <summary>
        /// The constant string for the Space
        /// </summary>
        private const string Space = @"\040";

        /// <summary>
        /// The constant string for the Tab
        /// </summary>
        private const string Tab = @"\t";

        /// <summary>
        /// The constant string for the open brackets
        /// </summary>
        private const string OpenBr = @"\[";

        /// <summary>
        /// The constant string for the close brackets
        /// </summary>
        private const string CloseBr = @"\]";

        /// <summary>
        /// The constant string for the open parentheses
        /// </summary>
        private const string OpenParen = @"\(";

        /// <summary>
        /// The constant string for the close parentheses
        /// </summary>
        private const string CloseParen = @"\)";

        /// <summary>
        /// The constant string for the Non-ASCII characters
        /// </summary>
        private const string NonAscii = @"\x80-\xff";

        /// <summary>
        /// The constant string for the Ctrl
        /// </summary>
        private const string Ctrl = @"\000-\037";

        /// <summary>
        /// The constant string for the carriage return/line feed
        /// </summary>
        private const string CRLF = @"\n\015";

        /// <summary>
        /// The regex expression for address
        /// </summary>
        private static Regex addreg;

        /// <summary>
        /// Initializes static members of the RFC822AddressParser class
        /// </summary>
        static RFC822AddressParser()
        {
            // Initialize the regex expression
            InitialRegex();
        }

        /// <summary>
        /// Verify whether the specified email address is compliant with RFC822 or not
        /// </summary>
        /// <param name="emailaddress">A string represent a actual email address</param>
        /// <returns>A value indicates whether the address is a valid email address, true if the specified emailaddress is compliant with RFC822, otherwise return false.</returns>
        public static bool IsValidAddress(string emailaddress)
        {
            return addreg.IsMatch(emailaddress);
        }

        /// <summary>
        /// Initialize the regex expression
        /// </summary>
        private static void InitialRegex()
        {
            string qtext = @"[^" + RFC822AddressParser.Escape +
                RFC822AddressParser.NonAscii +
                RFC822AddressParser.CRLF + "\"]";
            string dtext = @"[^" + RFC822AddressParser.Escape +
                RFC822AddressParser.NonAscii +
                RFC822AddressParser.CRLF +
                RFC822AddressParser.OpenBr +
                RFC822AddressParser.CloseBr + "\"]";

            string quoted_pair = " " + RFC822AddressParser.Escape + " [^" + RFC822AddressParser.NonAscii + "] ";
            string ctext = @" [^" + RFC822AddressParser.Escape +
                RFC822AddressParser.NonAscii +
                RFC822AddressParser.CRLF + "()] ";

            // Nested quoted pairs
            string cnested = string.Empty;
            cnested += RFC822AddressParser.OpenParen;
            cnested += ctext + "*";
            cnested += "(?:" + quoted_pair + " " + ctext + "*)*";
            cnested += RFC822AddressParser.CloseParen;

            // A comment
            string comment = string.Empty;
            comment += RFC822AddressParser.OpenParen;
            comment += ctext + "*";
            comment += "(?:";
            comment += "(?: " + quoted_pair + " | " + cnested + ")";
            comment += ctext + "*";
            comment += ")*";
            comment += RFC822AddressParser.CloseParen;

            // x is optional whitespace/comments
            string x = string.Empty;
            x += "[" + RFC822AddressParser.Space + RFC822AddressParser.Tab + "]*";
            x += "(?: " + comment + " [" + RFC822AddressParser.Space + RFC822AddressParser.Tab + "]* )*";

            // An email address atom
            string atom_char = @"[^(" + RFC822AddressParser.Space + ")<>\\@,;:\\\"." + RFC822AddressParser.Escape + RFC822AddressParser.OpenBr +
                RFC822AddressParser.CloseBr +
                RFC822AddressParser.Ctrl +
                RFC822AddressParser.NonAscii + "]";

            string atom = string.Empty;
            atom += atom_char + "+";
            atom += "(?!" + atom_char + ")";

            // Double quoted string, unrolled.
            string quoted_str = "(?'quotedstr'";
            quoted_str += "\\\"";
            quoted_str += qtext + " *";
            quoted_str += "(?: " + quoted_pair + qtext + " * )*";
            quoted_str += "\\\")";

            // A word is an atom or quoted string
            string word = string.Empty;
            word += "(?:";
            word += atom;
            word += "|";
            word += quoted_str;
            word += ")";

            // A domain-ref is just an atom
            string domain_ref = atom;

            // A domain-literal is like a quoted string, but [...] instead of "..."
            string domain_lit = string.Empty;
            domain_lit += RFC822AddressParser.OpenBr;
            domain_lit += "(?: " + dtext + " | " + quoted_pair + " )*";
            domain_lit += RFC822AddressParser.CloseBr;

            // A sub-domain is a domain-ref or a domain-literal
            string sub_domain = string.Empty;
            sub_domain += "(?:";
            sub_domain += domain_ref;
            sub_domain += "|";
            sub_domain += domain_lit;
            sub_domain += ")";
            sub_domain += x;

            // A domain is a list of subdomains separated by dots
            string domain = "(?'domain'";
            domain += sub_domain;
            domain += "(:?";
            domain += RFC822AddressParser.Period + " " + x + " " + sub_domain;
            domain += ")*)";

            // A route. A bunch of "@ domain" separated by commas, followed by a colon.
            string route = string.Empty;
            route += "\\@ " + x + " " + domain;
            route += "(?: , " + x + " \\@ " + x + " " + domain + ")*";
            route += ":";
            route += x;

            // A local-part is a bunch of 'word' separated by periods
            string local_part = "(?'localpart'";
            local_part += word + " " + x;
            local_part += "(?:";
            local_part += RFC822AddressParser.Period + " " + x + " " + word + " " + x;
            local_part += ")*)";

            // An addr-spec is local@domain
            string addr_spec = local_part + " \\@ " + x + " " + domain;

            // A route-addr is <route? addr-spec>
            string route_addr = string.Empty;
            route_addr += "< " + x;
            route_addr += "(?: " + route + " )?";
            route_addr += addr_spec;
            route_addr += ">";

            // A phrase
            string phrase_ctrl = @"\000-\010\012-\037";

            // Like atom-char, but without listing space, and uses phrase_ctrl. Since the class is negated, this matches the same as atom-char plus space and tab
            string phrase_char = "[^()<>\\@,;:\\\"." + RFC822AddressParser.Escape +
                RFC822AddressParser.OpenBr +
                RFC822AddressParser.CloseBr +
                RFC822AddressParser.NonAscii +
                phrase_ctrl + "]";

            string phrase = string.Empty;
            phrase += word;
            phrase += phrase_char;
            phrase += "(?:";
            phrase += "(?: " + comment + " | " + quoted_str + " )";
            phrase += phrase_char + " *";
            phrase += ")*";

            // A mailbox is an addr_spec or a phrase/route_addr
            string mailbox = string.Empty;
            mailbox += x;
            mailbox += "(?'mailbox'";
            mailbox += addr_spec;
            mailbox += "|";
            mailbox += phrase + " " + route_addr;
            mailbox += ")";

            RFC822AddressParser.addreg = new Regex(mailbox, RegexOptions.Compiled | RegexOptions.IgnorePatternWhitespace);
        }
    }
}
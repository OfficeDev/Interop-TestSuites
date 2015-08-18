namespace Microsoft.Protocols.TestSuites.MS_OXCFXICS
{
    using System.Collections.Generic;

    /// <summary>
    /// Supply help functions for manipulate Markers.
    /// </summary>
    public class MarkersHelper : EnumHelper
    {
        /// <summary>
        /// The length of a PidTag.
        /// </summary>
        public const int PidTagLength = 4;

        /// <summary>
        /// Indicate whether a Markers is an end marker.
        /// </summary>
        /// <param name="marker">The Markers.</param>
        /// <returns>If a Markers is an end marker, return true, else false.</returns>
        public static bool IsEndMarker(Markers marker)
        {
            List<Markers> markers = GetEnumValues<Markers>();
            return markers.Contains(marker);
        }

        /// <summary>
        /// Indicate whether a MetaProperties is an end marker.
        /// </summary>
        /// <param name="marker">The Markers.</param>
        /// <returns>If a Markers is an end marker, return true, else false.</returns>
        public static bool IsEndMarker(MetaProperties marker)
        {
            List<MetaProperties> markers = GetEnumValues<MetaProperties>();
            return markers.Contains(marker);
        }

        /// <summary>
        /// Indicate whether a value is an end marker.
        /// </summary>
        /// <param name="marker">A uint value.</param>
        /// <returns>If a Markers is an end marker, return true, else false.</returns>
        public static bool IsEndMarker(uint marker)
        {
            return IsEndMarker((Markers)marker)
            || IsEndMarker((MetaProperties)marker);
        }

        /// <summary>
        /// Indicate whether a value is an end marker.
        /// </summary>
        /// <param name="marker">A uint value.</param>
        /// <returns>If a Markers is an end marker, return true, else false.</returns>
        public static bool IsEndMarkerExceptEcWarning(uint marker)
        {
            if (marker != (uint)MetaProperties.PidTagEcWarning)
            {
                return IsEndMarker((Markers)marker)
|| IsEndMarker((MetaProperties)marker);
            }
            else
            {
                return IsEndMarker((Markers)marker);
            }
        }
    }
}
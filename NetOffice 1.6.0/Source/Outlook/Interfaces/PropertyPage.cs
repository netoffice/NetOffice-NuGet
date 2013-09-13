using System;
using NetRuntimeSystem = System;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Reflection;
using System.Collections.Generic;
using NetOffice;
namespace NetOffice.OutlookApi
{
	///<summary>
    /// Interface PropertyPage SupportByVersionAttribute Outlook, 9,10,11,12,14,15
    ///</summary>
    [SupportByVersionAttribute("Outlook", 9, 10, 11, 12, 14, 15)]
    [ComImport, Guid("0006307E-0000-0000-C000-000000000046"), TypeLibType(4096)]
    [EntityTypeAttribute(EntityType.IsInterface)]
    public interface PropertyPage
    {
        ///<summary>
        /// SupportByVersionAttribute Outlook, 9,10,11,12,14,15
        ///</summary>
        [SupportByVersionAttribute("Outlook", 9, 10, 11, 12, 14, 15)]
        [MethodImpl(4096)]
        void GetPageInfo([MarshalAs(19)] [In] [Out] ref string HelpFile, [In] [Out] ref int HelpContext);

        ///<summary>
        /// SupportByVersionAttribute Outlook, 9,10,11,12,14,15
        ///</summary>
        [SupportByVersionAttribute("Outlook", 9, 10, 11, 12, 14, 15)]
        [DispId(8449)]
        bool Dirty { [MethodImpl(4096)]get; }

        ///<summary>
        /// SupportByVersionAttribute Outlook, 9,10,11,12,14,15
        ///</summary>
        [SupportByVersionAttribute("Outlook", 9, 10, 11, 12, 14, 15)]
        [MethodImpl(4096)]
        void Apply();
    }
}
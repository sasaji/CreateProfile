using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace CreateProfile
{
    class Program
    {
        [DllImport("mapi32.dll")]
        [return: MarshalAs(UnmanagedType.Error)]
        internal static extern Int32 MAPIInitialize(IntPtr lpMapiInit);

        [DllImport("mapi32.dll")]
        [return: MarshalAs(UnmanagedType.Error)]
        internal static extern Int32 MAPIUninitialize();

        [return: MarshalAs(UnmanagedType.Error)]
        [DllImport("mapi32.dll")]
        internal static extern Int32 MAPIAdminProfiles(Int32 i, out IProfAdmin a);

        [return: MarshalAs(UnmanagedType.Error)]
        [DllImport("mapi32.dll")]
        internal static extern Int32 HrQueryAllRows(IMAPITable iMapiTable, IntPtr ptaga, IntPtr pres, IntPtr psos, long crowsMax, out IntPtr pprows);

        [Guid("00020301-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMAPITable
        {
            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 GetLastError([In][MarshalAs(UnmanagedType.Error)]Int32 hResult,
            [In][MarshalAs(UnmanagedType.U4)] UInt32 ulFlags,
            [Out][MarshalAs(UnmanagedType.LPStruct)]IntPtr lppMAPIError);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 Advise([In][MarshalAs(UnmanagedType.U4)]UInt32 ulEventMask,
            IntPtr lpAdviseSink,
            IntPtr lpulConnection);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 Advise([In][MarshalAs(UnmanagedType.U4)]UInt32 ulConnection);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 GetStatus(IntPtr lpulTableStatus,
            IntPtr lpulTableType);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 SetColumns(
            out IntPtr lpPropTagArray,
            [In][MarshalAs(UnmanagedType.U4)] UInt32 ulFlags);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 QueryColumns([In][MarshalAs(UnmanagedType.U4)]UInt32 ulFlags,
            IntPtr lpPropTagArray);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 GetRowCount([In][MarshalAs(UnmanagedType.U4)]UInt32 ulFlags,
            out UInt32 lpulCount);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 SeekRow(IntPtr bkOrigin,
            [In][MarshalAs(UnmanagedType.U4)]UInt32 lRowCount,
            IntPtr lplRowsSought);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 SeekRowApprox([In][MarshalAs(UnmanagedType.U4)]UInt32 ulNumerator,
            [In][MarshalAs(UnmanagedType.U4)]UInt32 ulDenominator);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 QueryPosition(IntPtr lpulRow,
            IntPtr lpulNumerator,
            IntPtr lpulDenominator);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 FindRow(
            out IntPtr lpRestriction,
            UInt32 BkOrigin,
            [In][MarshalAs(UnmanagedType.U4)]UInt32 ulFlags);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 Restrict(
            out IntPtr lpRestriction,
            [In][MarshalAs(UnmanagedType.U4)]UInt32 ulFlags);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 CreateBookmark(IntPtr lpbkPosition);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 FreeBookmark(IntPtr bkPosition);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 SortTable(IntPtr lpSortCriteria,
            [In][MarshalAs(UnmanagedType.U4)]UInt32 ulFlags);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 QuerySortOrder(IntPtr lppSortCriteria);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 QueryRows(
            UInt32 lRowCount,
            UInt32 ulFlags,
            out IntPtr lppRows);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 Abort();

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 ExpandRow([In][MarshalAs(UnmanagedType.U4)]UInt32 cbInstanceKey,
            IntPtr pbInstanceKey,
            [In][MarshalAs(UnmanagedType.U4)]UInt32 ulRowCount,
            [In][MarshalAs(UnmanagedType.U4)]UInt32 ulFlags,
            IntPtr lppRows,
            IntPtr lpulMoreRows);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 CollapseRow([In][MarshalAs(UnmanagedType.U4)]UInt32 cbInstanceKey,
            IntPtr pbInstanceKey,
            [In][MarshalAs(UnmanagedType.U4)]UInt32 ulFlags,
            IntPtr lpulRowCount);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 WaitForCompletion([In][MarshalAs(UnmanagedType.U4)]UInt32 ulFlags,
            [In][MarshalAs(UnmanagedType.U4)]UInt32 ulTimeout,
            IntPtr lpulTableStatus);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 GetCollapseState([In][MarshalAs(UnmanagedType.U4)]UInt32 ulFlags,
            [In][MarshalAs(UnmanagedType.U4)]UInt32 cbInstanceKey,
            IntPtr lpbInstanceKey,
            IntPtr lpcbCollapseState,
            IntPtr lppbCollapseState);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 SetCollapseState([In][MarshalAs(UnmanagedType.U4)]UInt32 ulFlags,
            [In][MarshalAs(UnmanagedType.U4)]UInt32 cbCollapseState,
            IntPtr pbCollapseState,
            IntPtr lpbkLocation);
        }

        [Guid("0002031c-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IProfAdmin
        {
            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 GetLastError([In][MarshalAs(UnmanagedType.Error)]Int32 hResult,
            [In][MarshalAs(UnmanagedType.U4)] UInt32 ulFlags,
            [Out][MarshalAs(UnmanagedType.LPStruct)]IntPtr lppMAPIError);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 GetProfileTable([In][MarshalAs(UnmanagedType.U4)] UInt32 ulFlags, [Out] out IMAPITable lppTable);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 CreateProfile(IntPtr lpszProfileName, IntPtr lpszPassword, IntPtr ulUIParam, UInt32 ulFlags);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 DeleteProfile(IntPtr lpszProfileName, UInt32 ulFlags);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 ChangeProfilePassword(
            [In][MarshalAs(UnmanagedType.LPWStr)]string lpszProfileName,
            [In][MarshalAs(UnmanagedType.LPWStr)] string lpszOldPassword,
            [In][MarshalAs(UnmanagedType.LPWStr)] string lpszNewPassword,
            [In][MarshalAs(UnmanagedType.U4)] UInt32 ulFlags);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 CopyProfile(
            IntPtr lpszOldProfileName,
            IntPtr lpszOldPassword,
            IntPtr lpszNewProfileName,
            IntPtr ulUIParam,
            IntPtr ulFlags);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 RenameProfile(
            IntPtr lpszOldProfileName,
            IntPtr lpszOldPassword,
            IntPtr lpszNewProfileName,
            IntPtr ulUIParam,
            IntPtr ulFlags);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 SetDefaultProfile(
            IntPtr lpszProfileName,
            UInt32 ulFlags);

            /*
            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 AdminServices(
            [In][MarshalAs(UnmanagedType.LPWStr)]string lpszProfileName,
            [In][MarshalAs(UnmanagedType.LPWStr)]string lpszPassword,
            [In][MarshalAs(UnmanagedType.U4)] UInt32 ulUIParam,
            [In][MarshalAs(UnmanagedType.U4)] UInt32 ulFlags,
            [Out][MarshalAs(UnmanagedType.IUnknown)] IntPtr lppServiceAdmin);
            */
            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 AdminServices(IntPtr lpszProfileName, IntPtr lpszPassword, UInt32 ulUIParam, UInt32 ulFlags, out IMsgServiceAdmin lppServiceAdmin);
        }

        [Guid("0002031d-0000-0000-c000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IMsgServiceAdmin
        {
            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 GetLastError([In][MarshalAs(UnmanagedType.Error)]Int32 hResult,
            [In][MarshalAs(UnmanagedType.U4)] UInt32 ulFlags,
            [Out][MarshalAs(UnmanagedType.LPStruct)]IntPtr lppMAPIError);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 GetMsgServiceTable(UInt32 ulFlags, out IMAPITable lppTable);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 CreateMsgService(IntPtr lpszService, IntPtr lpszDisplayName, IntPtr ulUIParam, UInt32 ulFlags);

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 DeleteMsgService();

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 CopyMsgService();

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 RenameMsgService();

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 ConfigureMsgService();

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 OpenProfileSection();

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 MsgServiceTransportOrder();

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 AdminProviders();

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 SetPrimaryIdentity();

            [PreserveSig]
            [return: MarshalAs(UnmanagedType.Error)]
            Int32 GetProviderTable();
        }

        public interface IProfSect
        {
        }

        struct SPropValue
        {
            public ulong ulPropTag;
            public ulong dwAlignPad;
            public string lpszA;
        }

        struct SRestriction
        {
            public ulong rt;
            public SContentRestriction resContent;
        }

        struct SContentRestriction
        {
            public ulong ulFuzzyLevel;
            public ulong ulPropTag;
            public IntPtr lpProp;
        }

        static int Main(string[] args)
        {
            uint PT_UNICODE = 31;
            uint PT_TSTRING = PT_UNICODE;
            uint PT_BINARY = 258;
            uint FL_FULLSTRING = 0x00000000;
            uint RES_CONTENT = 0x00000003;
            uint pidProfileMin = 0x6600;
            uint PR_DISPLAY_NAME_W = PROP_TAG(PT_UNICODE, 0x3001);
            uint PR_PROFILE_USER_SMTP_EMAIL_ADDRESS = PROP_TAG(PT_TSTRING, pidProfileMin + 0x41);
            uint PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W = PROP_TAG(PT_UNICODE, pidProfileMin + 0x41);
            uint PR_EMSMDB_SECTION_UID = PROP_TAG(PT_BINARY, 0x3D15);
            uint PR_SERVICE_NAME = PROP_TAG(PT_TSTRING, 0x3D09);
            uint MAPI_FORCE_ACCESS = 0x80000;

            Dictionary<uint, string> PropValueMap = new Dictionary<uint, string>() {
                { PR_DISPLAY_NAME_W, "Anthony" },
                { PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W, "SMTP:anthony@contoso.com" }
            };

            int hRes = 0;
            IProfAdmin lpProfAdmin;
            IMsgServiceAdmin lpSvcAdmin;
            IMAPITable lpMsgSvcTable;
            IntPtr lpSvcRows;
            List<SPropValue> rgvals = new List<SPropValue>();
            SRestriction sres = new SRestriction();
            SPropValue SvcProps = new SPropValue();
            IProfSect lpGlobalProfSection;
            IProfSect lpEmsMdbVarProfSect;
            SPropValue spvSmtpAddressW = new SPropValue();
            SPropValue spvDisplayName = new SPropValue();

            //enum { iSvcname, iSvcUID, cptaSvc };
            //SizedSPropTagArray(cptaSvc, sptCols) = { cptaSvc, PR_SERVICE_NAME, PR_SERVICE_UID };

            string profileName = "MAPIProfile";

            hRes = MAPIInitialize(IntPtr.Zero);
            if (hRes != 0) {
                Console.WriteLine("Error initializing MAPI.");
                goto error;
            }

            hRes = MAPIAdminProfiles(
                0,
                out lpProfAdmin);   // Pointer to new IProfAdmin.
            if (hRes != 0) {
                Console.WriteLine("Error getting IProfAdmin interface.");
                goto error;
            }

            hRes = lpProfAdmin.CreateProfile(
                Marshal.StringToHGlobalAnsi(profileName),   // Name of new profile.
                IntPtr.Zero,                                // Password for profile.
                IntPtr.Zero,                                // Handle to parent window.
                0);                                         // Flags.
            if (hRes != 0) {
                Console.WriteLine("Error creating profile.");
                goto error;
            }

            hRes = lpProfAdmin.AdminServices(
                Marshal.StringToHGlobalAnsi(profileName),   // Profile that we want to modify.
                IntPtr.Zero,                                // Password for that profile.
                0,                                          // Handle to parent window.
                0,                                          // Flags.
                out lpSvcAdmin);                            // Pointer to new IMsgServiceAdmin.
            if (hRes != 0) {
                Console.WriteLine("Error getting IMsgServiceAdmin interface.");
                goto error;
            }

            hRes = lpSvcAdmin.CreateMsgService(
                Marshal.StringToHGlobalAnsi("MSEMS"),   // Name of service from MAPISVC.INF.
                IntPtr.Zero,                            // Display name of service.
                IntPtr.Zero,                            // Handle to parent window.
                0);                                     // Flags.
            if (hRes != 0) {
                Console.WriteLine("Error creating Exchange message service.");
                goto error;
            }

            hRes = lpSvcAdmin.GetMsgServiceTable(
                0,                  // Flags.
                out lpMsgSvcTable); // Pointer to table.
            if (hRes != 0) {
                Console.WriteLine("Error getting Message Service Table.");
                goto error;
            }

            sres.rt = RES_CONTENT;
            sres.resContent.ulFuzzyLevel = FL_FULLSTRING;
            sres.resContent.ulPropTag = PR_SERVICE_NAME;
            IntPtr svcPropsPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(SvcProps));
            sres.resContent.lpProp = svcPropsPtr;

            SvcProps.ulPropTag = PR_SERVICE_NAME;
            SvcProps.lpszA = "MSEMS";
            Marshal.StructureToPtr(SvcProps, svcPropsPtr, true);
            IntPtr sresPtr = Marshal.AllocCoTaskMem(Marshal.SizeOf(sres));
            Marshal.StructureToPtr(sres, sresPtr, true);

            hRes = HrQueryAllRows(
                lpMsgSvcTable,
                IntPtr.Zero,
                sresPtr,
                IntPtr.Zero,
                0,
                out lpSvcRows);
            if (hRes != 0) {
                Console.WriteLine("Error querying table for new message service.");
                goto error;
            }

            Marshal.FreeCoTaskMem(svcPropsPtr);
            Marshal.FreeCoTaskMem(sresPtr);

        cleanup:
            MAPIUninitialize();
            return 0;
        error:
            Console.WriteLine(string.Format("hRes = {0}", hRes.ToString("X")));
            goto cleanup;
        }

        static uint PROP_TAG(uint ulPropType, uint ulPropID)
        {
            return (ulPropID << 16) | ulPropType;
        }
    }
}

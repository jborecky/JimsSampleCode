/*==========================================================================
' NAME: signed.exe
'
' AUTHOR: Jim Borecky , 
' DATE  : 5/4/2016
'
' COMMENTS: This binary is to return the certificate informatino of a file.
'
' ------------------------------------------------------------------
' Version   Date              Initial         Comment
' -------   --------          -------         -------
'  1.0      5/4/2016        Jim Borecky       Original
'  1.3      11/30/2017      Jim Borecky       Added serial number column
'  2.0      12/8/2017       Jim Borecky       Expanded to check catalog
'
'==========================================================================*/
using System;
using System.Security.Cryptography.X509Certificates;
using System.Runtime.InteropServices;
using CSCreateCabinet.Signature;
using Microsoft.Win32.SafeHandles;

namespace Signed
{
    class Program
    {
        static void Main(string[] args)
        {
            //Check for one parameter
            if(args.Length != 1)
            {
                Console.WriteLine("You need to supply a complete file name path.");
                return;
            }

            //Check for backslash
            if(args[0].IndexOf('\\') < 0)
            {
                Console.WriteLine("You need to supply a complete file name path.");
                return;
            }

            // Build the certificate structure
            X509Certificate2 theCertificate;
            bool InCatalog = false;

            //Attempt to build the certificate from the file. 
            //This will only happen if the cert is attached straight to the file.
            try
            {
                X509Certificate theSigner = X509Certificate.CreateFromSignedFile(args[0]);
                theCertificate = new X509Certificate2(theSigner);
            }
            catch (Exception ex)
            {
                //So we have no cert attached to the file. Time to check the catalogs
                if (IsInSignedCatalog(args[0]))
                {
                    InCatalog = true;
                }
                else
                {
                    Console.WriteLine("No digital signature found.");
                    return;
                }
            }

           if(InCatalog)
            {
                //Taking the catalog itself and see if it is signed.
                X509Certificate2 cataCertificate;
                try
                {                    
                    X509Certificate theSigner = X509Certificate.CreateFromSignedFile(Globals.catFile);
                    cataCertificate = new X509Certificate2(theSigner);
                    bool check = PrintCertificateInfo(cataCertificate, true);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error getting signature of catalog.");
                }
                
            }
            else
            {
                X509Certificate theSigner = X509Certificate.CreateFromSignedFile(args[0]);
                theCertificate = new X509Certificate2(theSigner);
                bool check = PrintCertificateInfo(theCertificate, false);
            }

        }
        //========================================================================================
        // static bool PrintCertificateInfo(X509Certificate2 theCertificate, bool InCatalog)
        //
        // Dependancies on other subs and Functions
        //		1) public static class Globals
        //
        // Comments:
        //	     This function prints out the information contained in the certificate
        //========================================================================================
        static bool PrintCertificateInfo(X509Certificate2 theCertificate, bool InCatalog)
        {
            /* This section will check that the certificate is from a trusted authority IE 
            * not self-signed. */
            try
            {
                bool chainIsValid = false;
                var theCertificateChain = new X509Chain();
                theCertificateChain.ChainPolicy.RevocationFlag = X509RevocationFlag.ExcludeRoot;
                theCertificateChain.ChainPolicy.RevocationMode = X509RevocationMode.Offline; //Checks to see if certificate is valid. Bad for the proxies
                theCertificateChain.ChainPolicy.UrlRetrievalTimeout = new TimeSpan(0, 1, 0);
                theCertificateChain.ChainPolicy.VerificationFlags = X509VerificationFlags.NoFlag;


                /*************************************************************************
                 * Do to the amount of time this part takes to process it has been removed.
                /*************************************************************************\
                //If the exe is in a catalog the chain is going to be self signed anyway.
                //Saving some execution time here.
                //if (!InCatalog)
                //    chainIsValid = theCertificateChain.Build(theCertificate);
                //else
                //    chainIsValid = false;
                \*************************************************************************/


                // In the event the attached cert is self-signed.
                if (chainIsValid)
                {
                    Console.WriteLine("Publisher Information: " + theCertificate.SubjectName.Name);
                    Console.WriteLine("Valid From: " + theCertificate.GetEffectiveDateString());
                    Console.WriteLine("Valid To: " + theCertificate.GetExpirationDateString());
                    Console.WriteLine("Issued By: " + theCertificate.Issuer);
                    Console.WriteLine("Serial Number: " + theCertificate.SerialNumber.ToString());
                    //Console.WriteLine("Chain Is Valid");
                }
                else
                {
                    Console.WriteLine("Publisher Information: " + theCertificate.SubjectName.Name);
                    Console.WriteLine("Valid From: " + theCertificate.GetEffectiveDateString());
                    Console.WriteLine("Valid To: " + theCertificate.GetExpirationDateString());
                    Console.WriteLine("Issued By: " + theCertificate.Issuer);
                    Console.WriteLine("Serial Number: " + theCertificate.SerialNumber.ToString());
                    if (InCatalog)
                    {
                        Console.WriteLine("In Catalog: " + Globals.catFile.ToString());
                    }
                    //Console.WriteLine("Chain Not Valid (certificate is self-signed)");
                }
                return true;
            }
            catch(Exception ex)
            {
                Console.WriteLine("Error occured getting certificate.");
                return false;
            }
        }

        //========================================================================================
        // static bool IsInSignedCatalog(string strFileName)
        //
        // Dependancies on other subs and Functions
        //		1) namespace CSCreateCabinet.Signature -- This has been modifed from the original 
        //              version to suit my needs.
        //      2) public static class Globals
        //
        // Comments:
        //	This ugly function checks to see if the hash value is part of the signed catalogs
        //========================================================================================
        static bool IsInSignedCatalog(string strFileName)
        {
            
            IntPtr Context;
            int HashSize = 0;
            byte[] emptyBuffer = new byte[] { 0 };
            IntPtr CatalogContext;
            bool ReturnFlag = false;

            /*Yes this seems to make no sense. These aren't drivers we are looking for but everything 
              is located in the C:\windows\System32\catroot\{F750E6C3-38EE-11D1-85E5-00C04FC295EE}
              directory which the DRIVER_ACTION_VERIFY guid matches. Go figure. */
            Guid guidAction = new Guid(NativeMethods.DRIVER_ACTION_VERIFY);

            //Get a context for signature verification.
            if (!NativeMethods.CryptCATAdminAcquireContext(out Context, guidAction, 0))
            {
               return false;
            }

            //Open a handle to the file that we are looking for.
            SafeFileHandle FileHandle = NativeMethods.CreateFileW(strFileName, NativeMethods.GENERIC_READ, NativeMethods.FILE_SHARE_READ, IntPtr.Zero, NativeMethods.OPEN_ALWAYS, 0, IntPtr.Zero);
            if (FileHandle.IsInvalid)
            {
                NativeMethods.CryptCATAdminReleaseContext(Context, 0);
                return false;
            }

            //Get the size we need for our hash using the file handle.
            NativeMethods.CryptCATAdminCalcHashFromFileHandle(FileHandle, ref HashSize, emptyBuffer, 0);
            if (HashSize == 0)
            {
                //0-sized has means error!
                NativeMethods.CryptCATAdminReleaseContext(Context, 0);
                NativeMethods.CloseHandle(FileHandle);
                return false;
            }

            //Allocate memory for the buffer.
            byte[] Buffer = new byte[HashSize];

            //Actually calculate the hash of the file
            if (!NativeMethods.CryptCATAdminCalcHashFromFileHandle(FileHandle, ref HashSize, Buffer, 0))
            {
                //Clean up if failed
                NativeMethods.CryptCATAdminReleaseContext(Context, 0);
                NativeMethods.CloseHandle(FileHandle);
                GC.Collect();
                return false;
            }

            //Get Catalog that the Hash is in, if any.
            CatalogContext = NativeMethods.CryptCATAdminEnumCatalogFromHash(Context, Buffer, Buffer.Length, 0, IntPtr.Zero);
            if (CatalogContext != IntPtr.Zero)
            {
                ReturnFlag = true;
            }
            else
            {
                ReturnFlag = false;
            }

            //Finish processing the catalog info here if any
            if (ReturnFlag)
            {
                //Get Catalog Name from Catalog
                NativeMethods.CATALOG_INFO hCatalogInfo = new NativeMethods.CATALOG_INFO();
                if (NativeMethods.CryptCATCatalogInfoFromContext(CatalogContext, out hCatalogInfo, 0))
                {
                    Globals.catFile = hCatalogInfo.wszCatalogFile;
                };
            }

            //Free context.
            if (CatalogContext != null)
                NativeMethods.CryptCATAdminReleaseCatalogContext(Context, CatalogContext, 0);
            NativeMethods.CryptCATAdminReleaseContext(Context, 0);
            GC.Collect();

            return ReturnFlag;
        }
    }
    //========================================================================================
    // public static class Globals
    //
    // Dependancies on other subs and Functions
    //		1) None
    //
    // Comments: Yeah dont judge me.
    //	       
    //========================================================================================
    public static class Globals
    {
        public static string catFile;
    }
}
//========================================================================================
// Code Beyond here, is my attempt to convert from c++ to C#
//========================================================================================
/*
 *  So YEAH, I butchered this code in order to make it work for me in C#.
 *  I added a lot of DLL functions and changed tons of variable types.
 *
 *  I could have cleaned out a lot of the summary's and structures since I 
 *  didn't use it all, but figured I'd hold onto it, in case I needed it for
 *  future code modifications.
 *
\***************************************************************************/
/****************************** Module Header ******************************\
 * Module Name:  NativeMethods.cs
 * Project:      CSCreateCabinet
 * Copyright (c) Microsoft Corporation.
 * 
 * This class wraps the extern method WinVerifyTrust in Wintrust.dll.
 *  
 * 
 * This source is subject to the Microsoft Public License.
 * See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
 * All other rights reserved.
 * 
 * THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, 
 * EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED 
 * WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.
\***************************************************************************/
namespace CSCreateCabinet.Signature
{
    internal static class NativeMethods
    {
        /// <summary>
        /// There is no interactive user. The trust provider performs the verification
        /// action without the user's assistance.
        /// </summary>
        public static readonly IntPtr INVALID_HANDLE_VALUE = new IntPtr(-1);

        /// <summary>
        ///  GUID of the action to verify a file or object using the Authenticode
        ///  policy provider.
        /// </summary>
        /// <summary>
        /// The WinVerifyTrust function performs a trust verification action on a 
        /// specified object. The function passes the inquiry to a trust provider 
        /// that supports the action identifier, if one exists.
        /// 
        /// INVALID_HANDLE_VALUE
        /// Zero
        /// A valid window handle
        /// </summary>
        /// <param name="hwnd">
        /// Optional handle to a caller window. A trust provider can use this value
        /// to determine whether it can interact with the user. However,
        /// trust providers typically perform verification actions without input from the user.
        /// 
        /// INVALID_HANDLE_VALUE
        /// Zero
        /// A valid window handle
        /// 
        /// </param>
        /// <param name="pgActionID">
        /// A pointer to a GUID structure that identifies an action and the trust
        /// provider that supports that action. This value indicates the type of 
        /// verification action to be performed on the structure pointed to by pWinTrustData.
        /// 
        /// DRIVER_ACTION_VERIFY
        /// HTTPSPROV_ACTION
        /// OFFICESIGN_ACTION_VERIFY
        /// WINTRUST_ACTION_GENERIC_CERT_VERIFY
        /// WINTRUST_ACTION_GENERIC_CHAIN_VERIFY
        /// WINTRUST_ACTION_GENERIC_VERIFY_V2
        /// WINTRUST_ACTION_TRUSTPROVIDER_TEST
        /// </param>
        public const string WINTRUST_ACTION_GENERIC_VERIFY_V2 =
                "{00AAC56B-CD44-11d0-8CC2-00C04FC295EE}";

        public const string WINTRUST_ACTION_GENERIC_CERT_VERIFY =
                "{189A3842-3041-11d1-85E1-00C04FC295EE}";

        public const string WINTRUST_ACTION_GENERIC_CHAIN_VERIFY =
                "{fc451c16-ac75-11d1-b4b8-00c04fb66ea0}";

        public const string DRIVER_ACTION_VERIFY =
                "{F750E6C3-38EE-11d1-85E5-00C04FC295EE}";
        /// <param name="pWVTData">
        /// A pointer that, when cast as a WINTRUST_DATA structure, contains information that the
        /// trust provider needs to process the specified action identifier. 
        /// Typically, the structure includes information that identifies the object that the 
        /// trust provider must evaluate.
        /// </param>
        /// <returns>
        /// If the trust provider verifies that the subject is trusted for the specified action,
        /// the return value is zero. 
        /// No other value besides zero should be considered a successful return.
        /// 
        /// For example, a trust provider might indicate that the subject is not trusted, or is 
        /// trusted but with limitations or warnings. The return value can be a trust-provider-specific 
        /// value described in the documentation for an individual trust provider, or it can be one
        /// of the following error codes.
        /// 
        /// TRUST_E_SUBJECT_NOT_TRUSTED
        /// TRUST_E_PROVIDER_UNKNOWN
        /// TRUST_E_ACTION_UNKNOWN
        /// TRUST_E_SUBJECT_FORM_UNKNOWN
        /// </returns>
        /// 

        /// GENERIC_WRITE -> (0x40000000L)
        public const int GENERIC_WRITE = 1073741824;

        public const uint GENERIC_READ = 2147483648;

        /// FILE_SHARE_DELETE -> 0x00000004
        public const int FILE_SHARE_DELETE = 4;

        /// FILE_SHARE_WRITE -> 0x00000002
        public const int FILE_SHARE_WRITE = 2;

        /// FILE_SHARE_READ -> 0x00000001
        public const int FILE_SHARE_READ = 1;

        /// OPEN_ALWAYS -> 4
        public const int OPEN_ALWAYS = 4;
        public const int CREATE_NEW = 1;

        public const int OPEN_EXISTING = 3;

        //All the wintrust.dll functions that I intend to use.
        [DllImport("wintrust.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool CryptCATCatalogInfoFromContext(
            [In] IntPtr hCatInfo, //HCATINFO
            [Out] out CATALOG_INFO phCatalogInfo, //HCATADMIN*
            [In] int dwFlags
        );
        [DllImport("wintrust.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern int WinVerifyTrust(
             [In] IntPtr hwnd,
             [In] [MarshalAs(UnmanagedType.LPStruct)] Guid pgActionID,
             [In] ref WINTRUST_DATA pWVTData
        );
        [DllImport("wintrust.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr CryptCATAdminEnumCatalogFromHash(
            [In] IntPtr hCatAdmin, //HCATADMIN
            [In] byte[] pbHash, //*
            [In] int cbHash,
            [In] int dwFlags,
            [In] IntPtr phPrevCatInfo //HCATINFO *
        );
        [DllImport("wintrust.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool CryptCATAdminCalcHashFromFileHandle(
            [In] SafeFileHandle hFile,
            ref int pcbHash, //*
            [In] byte[] pbHash, //*
            [In] int dwFlags
        );
        [DllImport("wintrust.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool CryptCATAdminReleaseCatalogContext(
           [In] IntPtr hCatAdmin, //HCATADMIN
           [In] IntPtr hCatInfo, //HCATINFO
           [In] int dwFlags
        );
        [DllImport("wintrust.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool CryptCATAdminAcquireContext(
          [Out]  out IntPtr phCatAdmin, //HCATADMIN*
          [In]  [MarshalAs(UnmanagedType.LPStruct)] Guid pgSubsystem, //*
          [In]  int dwFlags
        );
        [DllImport("wintrust.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern bool CryptCATAdminReleaseContext(
          [In] IntPtr hCatAdmin, //HCATADMIN
          [In] int dwFlags
        );
        [DllImportAttribute("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true, EntryPoint = "CreateFile")]
        public static extern SafeFileHandle CreateFileW(
            [InAttribute()] [MarshalAsAttribute(UnmanagedType.LPWStr)] string lpFileName,
            uint dwDesiredAccess,
            uint dwShareMode,
            [InAttribute()] System.IntPtr lpSecurityAttributes,
            uint dwCreationDisposition,
            uint dwFlagsAndAttributes,
            [InAttribute()] System.IntPtr hTemplateFile
        );
        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool CloseHandle(SafeFileHandle hHandle);
        /// <summary>
        /// The WINTRUST_DATA structure is used when calling WinVerifyTrust to pass
        /// necessary information into the trust providers.
        /// </summary>
        [StructLayout(LayoutKind.Sequential)]
        public class WINTRUST_DATA : IDisposable
        {

            /// <summary>
            /// The size, in bytes, of this structure.
            /// </summary>
            public UInt32 cbStruct;

            /// <summary>
            /// A pointer to a data buffer used to pass policy-specific data to a policy
            /// provider. This member can be NULL.
            /// </summary>
            public IntPtr pPolicyCallbackData;

            /// <summary>
            /// A pointer to a data buffer used to pass subject interface package 
            /// (SIP)-specific data to a SIP provider. This member can be NULL.
            /// </summary>
            public IntPtr pSIPClientData;

            /// <summary>
            /// Specifies the kind of user interface (UI) to be used. 
            /// </summary>
            public WTDUIChoice dwUIChoice;

            /// <summary>
            /// Certificate revocation check options. 
            /// </summary>
            public WTDRevocationChecks fdwRevocationChecks;

            /// <summary>
            /// Specifies the union member to be used and, thus, the type of object 
            /// for which trust will be verified. 
            /// </summary>
            public WTDUnionChoice dwUnionChoice;

            /// <summary>
            /// A pointer to a WINTRUST_FILE_INFO structure.
            /// 
            /// The definition of this field is
            /// 
            ///   union {
            ///       struct WINTRUST_FILE_INFO_  *pFile;
            ///       struct WINTRUST_CATALOG_INFO_  *pCatalog;
            ///       struct WINTRUST_BLOB_INFO_  *pBlob;
            ///       struct WINTRUST_SGNR_INFO_  *pSgnr;
            ///       struct WINTRUST_CERT_INFO_  *pCert;
            ///   };
            ///   
            /// We only use the file in this sample.
            /// </summary>
            public IntPtr pFile;

            /// <summary>
            /// Specifies the action to be taken. 
            /// </summary>
            public WTDStateAction dwStateAction;

            /// <summary>
            /// A handle to the state data. The contents of this member depends on 
            /// the value of the dwStateAction member.
            /// </summary>
            public IntPtr hWVTStateData;

            /// <summary>
            /// Reserved for future use. Set to NULL.
            /// </summary>
            [MarshalAs(UnmanagedType.LPWStr)]
            public String pwszURLReference;

            /// <summary>
            /// DWORD value that specifies trust provider settings. 
            /// </summary>
            public WTDProvFlags dwProvFlags;

            /// <summary>
            /// A DWORD value that specifies the user interface context for the 
            /// WinVerifyTrust function. This causes the text in the Authenticode
            /// dialog box to match the action taken on the file.
            /// </summary>
            public WTDUIContext dwUIContext;

            // constructor for silent WinTrustDataChoice.File check
            public WINTRUST_DATA(String fileName)
            {
                this.cbStruct = (UInt32)Marshal.SizeOf(typeof(WINTRUST_DATA));
                this.pPolicyCallbackData = IntPtr.Zero;
                this.pSIPClientData = IntPtr.Zero;
                this.dwUIChoice = WTDUIChoice.WTD_UI_NONE;
                this.fdwRevocationChecks = WTDRevocationChecks.WTD_REVOKE_NONE;
                this.dwUnionChoice = WTDUnionChoice.WTD_CHOICE_FILE;
                this.dwStateAction = WTDStateAction.WTD_STATEACTION_VERIFY;
                this.hWVTStateData = IntPtr.Zero;
                this.pwszURLReference = null;
                this.dwProvFlags = WTDProvFlags.WTD_HASH_ONLY_FLAG;
                this.dwUIContext = WTDUIContext.WTD_UICONTEXT_EXECUTE;

                WINTRUST_FILE_INFO wtfiData = new WINTRUST_FILE_INFO(fileName);
                this.pFile = Marshal.AllocCoTaskMem(Marshal.SizeOf(typeof(WINTRUST_FILE_INFO)));
                Marshal.StructureToPtr(wtfiData, pFile, false);
            }


            ~WINTRUST_DATA()
            {
                Dispose();
            }

            public void Dispose()
            {
                if (this.pFile != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(pFile);
                }
                GC.SuppressFinalize(this);
            }
        }

        #region WinTrustData struct field enums and structs
        public enum WTDUIChoice : uint
        {

            /// <summary>
            /// Display all UI.
            /// </summary>
            WTD_UI_ALL = 1,

            /// <summary>
            /// Display no UI.
            /// </summary>
            WTD_UI_NONE = 2,

            /// <summary>
            /// Do not display any negative UI.
            /// </summary>
            WTD_UI_NOBAD = 3,

            /// <summary>
            /// Do not display any positive UI.
            /// </summary>
            WTD_UI_NOGOOD = 4
        }

        public enum WTDRevocationChecks : uint
        {

            /// <summary>
            /// No additional revocation checking will be done when the WTD_REVOKE_NONE
            /// flag is used in conjunction with the HTTPSPROV_ACTION value set in the 
            /// pgActionID parameter of the WinVerifyTrust function. To ensure the 
            /// WinVerifyTrust function does not attempt any network retrieval when 
            /// verifying code signatures, WTD_CACHE_ONLY_URL_RETRIEVAL must be set in 
            /// the dwProvFlags parameter.
            /// </summary>
            WTD_REVOKE_NONE = 0,

            /// <summary>
            /// Revocation checking will be done on the whole chain.
            /// </summary>
            WTD_REVOKE_WHOLECHAIN = 1
        }

        public enum WTDUnionChoice : uint
        {

            /// <summary>
            /// Use the file pointed to by pFile.
            /// </summary>
            WTD_CHOICE_FILE = 1,

            /// <summary>
            /// Use the catalog pointed to by pCatalog.
            /// </summary>
            WTD_CHOICE_CATALOG = 2,

            /// <summary>
            /// Use the BLOB pointed to by pBlob.
            /// </summary>
            WTD_CHOICE_BLOB = 3,

            /// <summary>
            /// Use the WINTRUST_SGNR_INFO structure pointed to by pSgnr.
            /// </summary>
            WTD_CHOICE_SIGNER = 4,

            /// <summary>
            /// Use the certificate pointed to by pCert.
            /// </summary>
            WTD_CHOICE_CERT = 5

        }

        public enum WTDStateAction : uint
        {

            /// <summary>
            /// Ignore the hWVTStateData member.
            /// </summary>
            WTD_STATEACTION_IGNORE = 0,

            /// <summary>
            /// Verify the trust of the object (typically a file) that is specified by
            /// the dwUnionChoice member. The hWVTStateData member will receive a handle
            /// to the state data. 
            /// This handle must be freed by specifying the WTD_STATEACTION_CLOSE action
            /// in a subsequent call.
            /// </summary>
            WTD_STATEACTION_VERIFY = 1,

            /// <summary>
            /// Free the hWVTStateData member previously allocated with the 
            /// WTD_STATEACTION_VERIFY action.
            /// This action must be specified for every use of the WTD_STATEACTION_VERIFY action.
            /// </summary>
            WTD_STATEACTION_CLOSE = 2,

            /// <summary>
            /// Write the catalog data to a WINTRUST_DATA structure and then cache that structure. 
            /// This action only applies when the dwUnionChoice member contains WTD_CHOICE_CATALOG.
            /// </summary>
            WTD_STATEACTION_AUTO_CACHE = 3,

            /// <summary>
            /// Flush any cached catalog data. This action only applies when the dwUnionChoice
            /// member contains WTD_CHOICE_CATALOG.
            /// </summary>
            WTD_STATEACTION_AUTO_CACHE_FLUSH = 4

        }

        [Flags]
        public enum WTDProvFlags : uint
        {
            /// <summary>
            /// The trust is verified in the same manner as implemented by 
            /// Internet Explorer 4.0.
            /// </summary>
            WTD_USE_IE4_TRUST_FLAG = 0x1,

            /// <summary>
            /// The Internet Explorer 4.0 chain functionality is not used.
            /// </summary>
            WTD_NO_IE4_CHAIN_FLAG = 0x2,

            /// <summary>
            /// The default verification of the policy provider, such as code
            /// signing for Authenticode, is not performed, and the certificate 
            /// is assumed valid for all usages.
            /// </summary>
            WTD_NO_POLICY_USAGE_FLAG = 0x4,

            /// <summary>
            /// Revocation checking is not performed.
            /// </summary>
            WTD_REVOCATION_CHECK_NONE = 0x10,

            /// <summary>
            /// Revocation checking is performed on the end certificate only.
            /// </summary>
            WTD_REVOCATION_CHECK_END_CERT = 0x20,

            /// <summary>
            /// Revocation checking is performed on the entire certificate chain.
            /// </summary>
            WTD_REVOCATION_CHECK_CHAIN = 0x40,

            /// <summary>
            /// Revocation checking is performed on the entire certificate chain,
            /// excluding the root certificate.
            /// </summary>
            WTD_REVOCATION_CHECK_CHAIN_EXCLUDE_ROOT = 0x80,

            /// <summary>
            /// Not supported.
            /// </summary>
            WTD_SAFER_FLAG = 0x100,

            /// <summary>
            /// Only the hash is verified.
            /// </summary>
            WTD_HASH_ONLY_FLAG = 0x200,

            /// <summary>
            /// The default operating system version checking is performed.
            /// This flag is only used for verifying catalog-signed files.
            /// </summary>
            WTD_USE_DEFAULT_OSVER_CHECK = 0x400,

            /// <summary>
            /// If this flag is not set, all time stamped signatures are considered
            /// valid forever. Setting this flag limits the valid lifetime of the
            /// signature to the lifetime of the signing certificate. This allows 
            /// time stamped signatures to expire.
            /// </summary>
            WTD_LIFETIME_SIGNING_FLAG = 0x800,

            /// <summary>
            /// Use only the local cache for revocation checks. Prevents revocation
            /// checks over the network. 
            /// </summary>
            WTD_CACHE_ONLY_URL_RETRIEVAL = 0x1000,

            /// <summary>
            /// Disable the use of MD2 and MD4 hashing algorithms. If a file is signed by 
            /// using MD2 or MD4 and if this flag is set, an NTE_BAD_ALGID error is returned.
            /// 
            /// Note:
            /// This flag is only supported on Windows 7 with SP1 and later operating systems.
            /// </summary>
            WTD_DISABLE_MD2_MD4 = 0x2000
        }

        public enum WTDUIContext : uint
        {

            /// <summary>
            /// Use when calling WinVerifyTrust for a file that is to be run. 
            /// This is the default value.
            /// </summary>
            WTD_UICONTEXT_EXECUTE = 0,

            /// <summary>
            /// Use when calling WinVerifyTrust for a file that is to be installed.
            /// </summary>
            WTD_UICONTEXT_INSTALL = 1

        }

        [StructLayout(LayoutKind.Sequential)]
        public struct WINTRUST_FILE_INFO
        {
            public uint cbStruct;

            [MarshalAs(UnmanagedType.LPWStr)]
            public string pcwszFilePath;

            public IntPtr hFile;

            public IntPtr pgKnownSubject;

            public WINTRUST_FILE_INFO(string filePath)
            {
                this.cbStruct = (uint)Marshal.SizeOf(typeof(WINTRUST_FILE_INFO));
                this.hFile = IntPtr.Zero;
                this.pgKnownSubject = IntPtr.Zero;
                this.pcwszFilePath = filePath;
            }
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct CATALOG_INFO
        {

           public UInt32 cbStruct;

           [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 256)]
           public string wszCatalogFile;
            #endregion

        }
    }
}



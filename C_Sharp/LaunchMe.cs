using System;
using System.Collections.Generic;
using System.Text;
using Rijndael;
using System.Diagnostics;
using System.Security;

namespace LaunchMe
{
    class Program
    {
        static void Main(string[] args)
        {
            //-----------------------------------------------------------------------
            //Encrypt and Decrypt
            string tempTest = "";
            string passPhrase = "cra55pr@ae";   //can be any string
            string saltValue = "s@ltJ3$u$";     //can be any string
            string hashAlgorithm = "SHA1";      //Can be MD5
            int passwordIterations = 5;         //can be any number
            string initVector = "@234C887v2335dd4"; //must be 16 bytes
            int keySize = 256;                  //can be 192 or 128
            string plainText = "";
            string filePara = "";
            string fileEXE = "";
            bool hashFlag = false;
            bool cmdFlag = false;
            string strEnteredAccount = "";
            string strEnteredPassword = "";
            string strDomain = "";
            
            if (args.Length > 0)
            {
                for (int i = 0; i < args.Length; i++)
                {
                    tempTest = args[i].ToLower();
                    if (tempTest.IndexOf("/account:") != -1)
                    {
                        //Pull account info here
                        plainText = args[i].Remove(0, 9);
                    }
                    else if (tempTest.IndexOf("/fileexe:") != -1)
                    {
                        //Pull specific domain here
                        fileEXE = args[i].Remove(0, 9);
                        fileEXE = fileEXE.ToLower();
                        cmdFlag = true;
                    }
                    else if (tempTest.IndexOf("/filepara:") != -1)
                    {
                        //Pull specific domain here
                        filePara = args[i].Remove(0, 10);
                        filePara = filePara.ToLower();
                        //Need to check for proper formatting.
                    }
                    else
                    {
                        //Sorting through args
                        switch (args[i].ToLower())
                        {
                            case "/hash":
                                hashFlag = true;
                                break;
                            default:
                                ShowHelp();
                                Environment.Exit(0);
                                break;
                        }
                    }

                    //Console.WriteLine(args[i]);
                }
                //process args here.
                if (hashFlag && cmdFlag)
                {
                    ShowHelp();
                    Environment.Exit(0);
                }

                if (hashFlag)
                {
                    //Need to check for proper formatting.
                    if ((plainText.IndexOf("\\") == -1) || (plainText.IndexOf(";") == -1))
                    {
                        ShowHelp();
                        Environment.Exit(0);
                    }

                    Console.WriteLine("Finding cipher for {0}", plainText);
                    string cipherText = RijndaelSimple.Encrypt(plainText,
                                                                passPhrase,
                                                                saltValue,
                                                                hashAlgorithm,
                                                                passwordIterations,
                                                                initVector,
                                                                keySize);
                    Console.WriteLine("Your Hash value is: {0}", cipherText);
                }
                else if ((plainText != ""))
                {
                    string decipherText = "";
                    try
                    {
                        decipherText = RijndaelSimple.Decrypt(plainText,
                                                            passPhrase,
                                                            saltValue,
                                                            hashAlgorithm,
                                                            passwordIterations,
                                                            initVector,
                                                            keySize);
                        //Console.WriteLine(String.Format("Decrypted : {0}", decipherText));
                    }
                    catch
                    {
                        Console.WriteLine("");
                        Console.WriteLine("");
                        Console.WriteLine("********** WARNING ERROR PROCESS ABORTED ************");
                        Console.WriteLine("     An error occured most likely a invalid hash");
                        Console.WriteLine("*****************************************************");
                        Environment.Exit(0);
                    }

                    //ok time to break up the pieces.

                    string[] tmpArray = decipherText.Split('\\');
                    strDomain = tmpArray[0];
                    string strAccountAndPassword = tmpArray[1];
                    tmpArray = strAccountAndPassword.Split(';');
                    strEnteredAccount = tmpArray[0];
                    strEnteredPassword = tmpArray[1];

                    //Console.WriteLine("Domain:    {0}", strDomain);
                    //Console.WriteLine("Account:   {0}", strEnteredAccount);
                    //Console.WriteLine("Password:  {0}", strEnteredPassword);
                    Console.WriteLine("Running Command:  {0}", fileEXE + " " + filePara);

                    //Launch the utility here.
                    Process WMT = new Process();
                    WMT.StartInfo.Domain = strDomain;
                    WMT.StartInfo.UserName = strEnteredAccount;
                    char[] arrPass = strEnteredPassword.ToCharArray();
                    SecureString myPassword = new SecureString();
                    foreach (char character in arrPass)
                    {
                        myPassword.AppendChar(character);
                    }
                    WMT.StartInfo.Password = myPassword;
                    WMT.StartInfo.FileName = fileEXE;
                    WMT.StartInfo.Arguments = filePara;
                    WMT.StartInfo.UseShellExecute = false;

                    WMT.Start();

                    WMT.WaitForExit();
                    int exitCode = WMT.ExitCode;

                    Environment.Exit(exitCode);
                }
                else
                {
                    ShowHelp();
                    Environment.Exit(0);
                }

            }
            else
            {
                ShowHelp();
            }
        }

        //==================================================================================
        // static void ShowHelp()
        //
        // Dependancies on other Subs or Functions
        //	1)None
        //
        //Comments:
        // 	This Function shows the help menu context
        //==================================================================================
        static void WriteEvent(string sEvent)
        {
            string sSource = "JBJoin";
            string sLog = "Application";

            if (!EventLog.SourceExists(sSource))
                EventLog.CreateEventSource(sSource, sLog);
            EventLog.WriteEntry(sSource, sEvent, EventLogEntryType.Error);
        }
        //==================================================================================
        // static void ShowHelp()
        //
        // Dependancies on other Subs or Functions
        //	1)None
        //
        //Comments:
        // 	This Function shows the help menu context
        //==================================================================================
        static void ShowHelp()
        {
            Console.WriteLine("===============================================================================");
            Console.WriteLine(":         (c) 2011");
            Console.WriteLine("Usage:  LaunchMe.exe (/hash || (/FileEXE:executable && [/FilePara:mystuff]))");
            Console.WriteLine("/account:domain\\myID;MyPassword ");
            Console.WriteLine("-------------------------------------------------------------------------------");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("");
            Console.WriteLine("/hash           This will return the hash value of the");
            Console.WriteLine("                account and password supplied.");
            Console.WriteLine("/FileEXE:       This is the executable file needed to run under alternate");
            Console.WriteLine("                credentials.");
            Console.WriteLine("/FilePara:      This is the parameter string needed to run the executable.");
            Console.WriteLine("/account:       Account and password must be properly formatted.");
            Console.WriteLine("                Example:  /account:rougeone\\meID;MyPassword");
            Console.WriteLine("                          /account:LHWLoAh77dP2qgXRRUHvf8IPVddl1zDCyLlwKQ5gSeA=");
            Console.WriteLine("===============================================================================");
        }
    }
}

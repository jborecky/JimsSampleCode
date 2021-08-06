/*==========================================================================
' NAME: signed.exe
'
' AUTHOR: Jim Borecky , 
' DATE  : 5/10/2016
'
' COMMENTS: Just getting the file attributes.
'
' ------------------------------------------------------------------
' Version   Date              Initial         Comment
' -------   --------          -------         -------
'  1.0      5/10/2016        Jim Borecky       Original
'
'==========================================================================*/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Diagnostics;

namespace FileDetails
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"C:\Users\a426971\Desktop";
            string fileName = @"JimTest.txt";

            // get file attributes
            FileAttributes fileAttributes = File.GetAttributes(filePath + "\\" + fileName);

            // clear all file attributes
            File.SetAttributes(filePath + "\\" + fileName, FileAttributes.Normal);

            FileVersionInfo.GetVersionInfo(Path.Combine(filePath, fileName));
            FileVersionInfo myFileVersionInfo = FileVersionInfo.GetVersionInfo(filePath + "\\" + fileName);
            
            
            // Print the file name and version number.
            Console.WriteLine("File Description: " + myFileVersionInfo.FileDescription + '\n' +
                              "Version number: " + myFileVersionInfo.FileVersion);

            myFileVersionInfo.ProductBuildPart = "1.1.1.1";
            myFileVersionInfo.

            //File.SetAttributes(filePath + "\\" + fileName, FileAttributes);

            Console.ReadLine();

        }
    }
}

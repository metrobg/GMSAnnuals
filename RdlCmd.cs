/* ====================================================================
    Copyright (C) 2004-2008  fyiReporting Software, LLC

    This file is an example showing using the fyiReporting RDL project.
	
    You may modify and use this file in any fashion you want.  The RdlEngine.dll
	module is available from fyiReporting Software, LLC and is licensed under 
	the Apache Version 2 license.  

    For additional information, email info@fyireporting.com or visit
    the website www.fyiReporting.com.
 * 
 * 
 * ward 4963  2 bank documents
 * ward 6407  1 bank document
 * wards 6447 & 6147 0 bank documents
 * 
 * 
 *  
 * /faccounting || plan -w6407 -tpdf  -pmerlin -s01/01/2014 -e12/31/2014 -uaccounting || -uplan
 * 
 * 
 *   * Push command line parameters -s -e and -w to report as report parameters GG
                * 
                * c:\work\RdlCmd\bin\Debug>rdlcmd /faccounting|| plan -tpdf  -oC:\work  -pmerlin -w5515  -uAccounting || -uPlan
                * -s01/01/2014 -e12/31/2014
 * 
 * 
 * 
*/

using System;
using System.IO;
using System.Net;
using System.Collections;
using System.Collections.Specialized;
using System.Configuration;
using System.Reflection;
using System.Drawing;
using System.Drawing.Imaging;
using System.Text.RegularExpressions;
using System.Globalization;
using fyiReporting.RDL;
using Oracle.DataAccess;
using Oracle.DataAccess.Client;
using ceTe.DynamicPDF;
using ceTe.DynamicPDF.Merger;

/**************************************************************************************************************************************************************************
                * Push command line parameters -s -e and -w to report as report parameters GG
                * 
                * c:\work\RdlCmd\bin\Debug>rdlcmd /fC:\Users\Auditor\Documents\work\Legal-AnnualAccountingRevised_2015.rdl?rc:ofile=5515 -tpdf  -oC:\work  -pmerlin -w5515  -uAccounting || -uPlan
                * -s01/01/2014 -e12/31/2014
                * ******************************************************************************************************************************************************* */
namespace fyiReporting.RdlCmd
{
    /// <summary>
    /// RdlCmd is a batch report generation program.  It takes a report definition
    ///   and renders it into the requested formats.
    /// </summary>
    public class RdlCmd
    {
        /// <summary>
        /// RdlCmd takes a report definition and renders it in the specified formats.
        /// </summary>

        [STAThread]
        static public int Main(string[] args)
        {
            //  development GPDC-PC60
            // ceTe.DynamicPDF.Document.AddLicense("MER80NSSNPPNFDJ/+tFdG64tA51eMVpcD/SLKaMyia0Lo3DoXqKwR0Q/TT2bCdTLvdqoDUzoGkaCO178Bvj6lABsHvxDUGcptI8g");

            // GPDC-ORACLE  (Production / 60.8 and Test 144.234)

            ceTe.DynamicPDF.Document.AddLicense("MER80NSSNPPNFDrknfU84ZvBJwN89tBRokt4dClNfa11CGFZi895ceFyCUFkb0dFv5qDno6eEM5tGmBz1+MZ1p4AjttBviN//ZFg");


            // Handle the arguments""
            if (args == null || args.Length == 0)
            {


                Console.WriteLine(string.Format("RdlCmd Version {0}, Copyright (C) 2004-2008 Advanced Consulting Enterprises",
                                Assembly.GetExecutingAssembly().GetName().Version.ToString()));
                Console.WriteLine("");
                Console.WriteLine("RdlCmd comes with ABSOLUTELY NO WARRANTY.  This is free software,");
                Console.WriteLine("and you are welcome to redistribute it under certain conditions.");
                Console.WriteLine("");
                Console.WriteLine("For help, type RdlCmd /?");
            }



            Ward ward;
            DocHelper dh = new DocHelper();
            ReportRunner rr = new ReportRunner();

            string[] files = new string[1];
            string[]  types = new string[1];

            char[] breakChars = new char[] { '+' };
           // string files = " ";
            string logFile = null;
            string dir = null;
            int cnt = 0;

            string outputFolder = null;
            string reportType = "accounting";
           

            string oradb = ConfigurationManager.AppSettings["dbconnection"];
            outputFolder = ConfigurationManager.AppSettings["AnnualAccounting"];   // default folder  p:\Annuals\AnnualAccounting or p:\Annuals\AnnualPlan
            logFile = ConfigurationManager.AppSettings["logFile"];

            string StartDate = null;
            string EndDate = null;
            string Ward = null;
            string reportURL = null;

            types[0] = "pdf";

            foreach (string s in args)
            {

                string t = s.Substring(0, 2);
                cnt++;

                switch (t)
                {
                    case "/e":
                    case "-e":
                        EndDate = s.Substring(2);
                        break;
                    case "/f":             // choices are -faccounting or -fplan or -fcjis
                    case "-f":
                        if (s.Substring(2).ToLower() == "cjis")
                        {
                            outputFolder = ConfigurationManager.AppSettings["CJISMemoFolder"];
                            reportType = "cjis";
                        }
                        if (s.Substring(2).ToLower() == "accounting")
                        {
                            files[0] = ConfigurationManager.AppSettings["AccountingURL"];
                            outputFolder = ConfigurationManager.AppSettings["AnnualAccountingFolder"];
                            reportType = "accounting";
                        }
                        if (s.Substring(2).ToLower() == "plan")
                        {
                            reportURL = ConfigurationManager.AppSettings["AnnualPlanURL"];
                            outputFolder = ConfigurationManager.AppSettings["AnnualPlanFolder"];
                            reportType = "plan";
                        }
                        break;
                    case "/s":
                    case "-s":
                        StartDate = s.Substring(2);
                        break;
                    case "-w":
                    case "/w":
                        Ward = s.Substring(2);
                        break;
                    default:
                        Console.WriteLine("Unknown command '{0}' ignored.", s);
                        //returnCode = 4;
                        break;
                }
            }
            if (files == null)
            {
                Console.WriteLine("/f parameter is required.");
                return 8;
            }

            if (dir == null)
            {
                //dir = Environment.CurrentDirectory;
                dir = outputFolder;
            }

            if (dir[dir.Length - 1] != Path.DirectorySeparatorChar)
                dir += Path.DirectorySeparatorChar;

            OracleConnection connection = new OracleConnection(oradb);
            connection.Open();
            Console.WriteLine("DB Connection is: {0}", connection.State.ToString());

            ward = dh.getWardInformation(int.Parse(Ward), connection);


            if (reportType == "accounting")
            {
                Console.WriteLine("Processing Annual Accounting");
                createAccountingPDF(rr, ward, dh, files, outputFolder, StartDate, EndDate, connection);
            }

            if (reportType == "plan")
            {
                createAnnualPlanPDF(ward, dh, reportURL, outputFolder, StartDate, EndDate, connection, logFile);                                                                     // end of if accounting
                Console.WriteLine("Processing Annual Plan");
            }
            if (reportType == "cjis")
            {
                createCJISPDF(ward, dh, outputFolder, connection, logFile);
                Console.WriteLine("Processed CJIS Memo");     // end of if CJIS             
            }
            return 1;

        }    /*   end of main  */

        private static void createAccountingPDF(ReportRunner rr, Ward ward, DocHelper dh, string[] files, string outputFolder, string StartDate, string EndDate, OracleConnection connection)
        {
            decimal wardDocCount = 0;
            WardDocument[] wardDocArray;
            string[] types = null;
            types = new string[] { "pdf" };
            string pdfTargetFolder = null;

            wardDocCount = dh.getBankStatementCount(ward.getWardNumber(), connection, EndDate);                         // count all of the documents we need, doctype "MELLON" and "BANK"
            wardDocArray = new WardDocument[(int)wardDocCount];
            pdfTargetFolder = System.IO.Path.Combine(outputFolder, ward.getWardName() + "_" + ward.getFileNumber());


            if (!System.IO.Directory.Exists(pdfTargetFolder))                                                           // create output directory as needed for each ward
            {
                System.IO.Directory.CreateDirectory(pdfTargetFolder);
            }

            string document1 = System.IO.Path.Combine(pdfTargetFolder, ward.getWardNumber() + "_temporary.pdf");        // the document created by the RDL component
            string document2;
            MergeDocument document;


            if (wardDocCount > 0)
            {

                dh.getWardBankDocs(wardDocArray, connection, ward.getWardNumber(), EndDate);                            // populate wardDocArray with the documents
                Console.WriteLine("Ward has {0} Bank document(s) to copy", wardDocCount);

                // rc.returnCode = returnCode;  // MBG 08/15/15
                rr.DoRender(pdfTargetFolder, files, types, ward.getWardNumber().ToString(), StartDate, EndDate);                                   //   create the AnnualAccounting PDF document
                document = new MergeDocument(document1);

                //   B E G I N       M E R G E   O F   P D F    D O C U M E N T S           

                Console.WriteLine("output folder is: {0}", pdfTargetFolder);
                // loop over any additional documents here and merge with the AnnualAccouting pdf

                foreach (WardDocument w in wardDocArray)
                {
                    if (File.Exists(System.IO.Path.Combine(w.getStoragePath(), w.getDocPath())))
                    {
                        document.Append(System.IO.Path.Combine(w.getStoragePath(), w.getDocPath()));
                    }
                    else
                    {
                        System.IO.File.AppendAllText(@"P:\Annuals\AnnualAccountingLog.txt", "Could not find ward document: " + w.getStoragePath() + "/" + w.getDocPath() + " - " + DateTime.Now + "\r\n");
                    }
                }


                ceTe.DynamicPDF.PageList pl = new PageList();                 // find out how many pages the resulting PDF document has so we can add the count to the final filename
                pl = document.Pages;
                Console.WriteLine("File has {0} Pages", pl.Count);
                document2 = System.IO.Path.Combine(pdfTargetFolder, ward.getWardNumber() + "_AnnualAccounting_" + pl.Count.ToString() + ".pdf"); // name of the final output document
                document.Draw(document2);
                File.Delete(document1);      // delete the 9999_temporary.pdf
                document2 = null;

            }
            else
            {
                Console.WriteLine("LOG: => No Bank Documents found for ward {0}", ward.getWardNumber());
                System.IO.File.AppendAllText(@"C:\Annuals\AnnualAccountingLog.txt", "Ward: " + ward.getWardNumber() + " has no Bank Statements to process " + DateTime.Now + "\r\n");

                // rc.returnCode = returnCode;  // MBG 08/15/15
                rr.DoRender(pdfTargetFolder, files, types, ward.getWardNumber().ToString(), StartDate, EndDate);
                ceTe.DynamicPDF.PageList pl = new PageList();
                document = new MergeDocument(document1);
                pl = document.Pages;
                document2 = System.IO.Path.Combine(pdfTargetFolder, ward.getWardNumber() + "_AnnualAccounting_" + pl.Count.ToString() + ".pdf"); // name of the final output document
                document.Draw(document2);
                File.Delete(document1);                                         // delete the 9999_temporary.pdf
                document2 = null;

                connection.Close();
            }
            dh.updateAnnualAccounting(ward.getWardNumber(), connection);
            connection.Close();
        }       /* end of createAccountingPDF  */

        private static void createAnnualPlanPDF(Ward ward, DocHelper dh, string reportURL, string outputFolder, string StartDate, string EndDate, OracleConnection connection, string logFile)
        {
            decimal physicianDocCount = 0;
            WardDocument[] wardDocArray;
            string[] types = null;
            types = new string[] { "pdf" };
            string pdfTargetFolder = null;

            Console.WriteLine("Processing Annual Plan");
            physicianDocCount = dh.getPhysicianDocumentCount(ward.getWardNumber(), connection);          // count all of the documents we need, doctype "MELLON" and "BANK"
            wardDocArray = new WardDocument[(int)physicianDocCount];

            pdfTargetFolder = System.IO.Path.Combine(outputFolder, ward.getWardName() + "_" + ward.getFileNumber());

            if (!System.IO.Directory.Exists(pdfTargetFolder))                                                           // create output directory as needed for each ward
            {
                System.IO.Directory.CreateDirectory(pdfTargetFolder);
            }

            string document1 = System.IO.Path.Combine(pdfTargetFolder, ward.getWardNumber() + "_deleteME.pdf");        // the document created by the RDL component
            string document2;
            MergeDocument document;

            /*    C R E A T E    T H E   B A S E  A N N U A L  P L A N   R E P O R T  */
            WebClient client = new WebClient();
            string url = reportURL + ward.getWardNumber(); ;
            client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
            client.DownloadFile(url, document1);

            if (physicianDocCount > 0)
            {

                dh.getWardPhysicianReport(wardDocArray, connection, ward.getWardNumber());                            // populate wardDocArray with the documents
                Console.WriteLine("Ward has {0} Physician Report(s) to copy", physicianDocCount);
                dh.updatePlanStatus(ward.getWardNumber(), connection);
                connection.Close();
                //   create the AnnualAccounting PDF document
                document = new MergeDocument(document1);



                //   B E G I N       M E R G E   O F   P D F    D O C U M E N T S           

                Console.WriteLine("output folder is: {0}", pdfTargetFolder);
                // loop over any additional documents here and merge with the AnnualAccouting pdf

                foreach (WardDocument w in wardDocArray)
                {
                    if (File.Exists(System.IO.Path.Combine(w.getStoragePath(), w.getDocPath())))
                    {
                        document.Append(System.IO.Path.Combine(w.getStoragePath(), w.getDocPath()));
                    }
                    else
                    {
                        System.IO.File.AppendAllText(@"AnnualAccountingLog.txt", "Could not find ward document: " + w.getStoragePath() + "\\" + w.getDocPath() + " - " + DateTime.Now + "\r\n");
                    }
                }


                ceTe.DynamicPDF.PageList pl = new PageList();                 // find out how many pages the resulting PDF document has so we can add the count to the final filename
                pl = document.Pages;
                Console.WriteLine("File has {0} Pages", pl.Count);
                document2 = System.IO.Path.Combine(pdfTargetFolder, ward.getWardNumber() + "_AnnualPlan_" + pl.Count.ToString() + ".pdf"); // name of the final output document
                document.Draw(document2);
                File.Delete(document1);      // delete the 9999_temporary.pdf
                document2 = null;

            }
            else
            {
                Console.WriteLine("LOG: => No Physician Documents found for ward {0}", ward.getWardNumber());
                System.IO.File.AppendAllText(@logFile, "Ward: " + ward.getWardNumber() + " has no Doctor Reports to process " + DateTime.Now + "\r\n");

                ceTe.DynamicPDF.PageList pl = new PageList();
                document = new MergeDocument(document1);
                pl = document.Pages;
                document2 = System.IO.Path.Combine(pdfTargetFolder, ward.getWardNumber() + "_AnnualPlan_" + pl.Count.ToString() + ".pdf"); // name of the final output document
                document.Draw(document2);
                File.Delete(document1);                                         // delete the 9999_temporary.pdf
                document2 = null;
                dh.updatePlanStatus(ward.getWardNumber(), connection);
                connection.Close();
            }
        }

        private static void createCJISPDF(Ward ward, DocHelper dh, string outputFolder, OracleConnection connection, string logFile)
        {
            decimal guardianLetterCount = 0;
            decimal cjisMemoCount = 0;
            WardDocument[] wardDocArray;
            //string[] types = null;
            //types = new string[] { "pdf" };
            string pdfTargetFolder = null;

            PdfDocument profilePDF = null;
            PdfDocument memoPDF = null;
            PdfDocument authPDF = null;
            MergeDocument document = new MergeDocument();

            string cjisProfileURL = ConfigurationManager.AppSettings["CJIS_ProfileURL"];
            string cjisMemoURL = ConfigurationManager.AppSettings["CJIS_MemoURL"];
            string cjisAuthURL = ConfigurationManager.AppSettings["CJIS_AuthURL"];

            guardianLetterCount = dh.getguardianLetterCount(ward.getWardNumber(), connection);          // count all of the documents we need, doctype "GRDLET"

            wardDocArray = new WardDocument[(int)guardianLetterCount];             // create an array large enough to contain all of the necessary documents

            pdfTargetFolder = System.IO.Path.Combine(outputFolder, ward.getWardName() + "_" + ward.getWardNumber());

            if (!System.IO.Directory.Exists(pdfTargetFolder))                                                           // create output directory as needed for each ward
            {
                System.IO.Directory.CreateDirectory(pdfTargetFolder);
            }


            string profileDocumentPath = System.IO.Path.Combine(pdfTargetFolder, ward.getWardNumber() + "_Profile.pdf");        // the document created by Jasper
            string memoDocumentPath = System.IO.Path.Combine(pdfTargetFolder, ward.getWardNumber() + "_Memo.pdf");
            string authDocumentPath = System.IO.Path.Combine(pdfTargetFolder, ward.getWardNumber() + "_Auth.pdf");
            string document2;

            /*    C R E A T E    T H E  3  B A S E  C J I S    R E P O R T S */
            WebClient client = new WebClient();
            string url = cjisProfileURL + ward.getWardNumber();
            client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
            client.DownloadFile(url, profileDocumentPath);

            url = cjisMemoURL + ward.getWardNumber();
            client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
            client.DownloadFile(url, memoDocumentPath);

            url = cjisAuthURL + ward.getWardNumber();
            client.Headers.Add("user-agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)");
            client.DownloadFile(url, authDocumentPath);

            if (File.Exists(memoDocumentPath))
            {            // make the downloaded documents pdf objects and prepare to merge them
                cjisMemoCount = 1;
                profilePDF = new PdfDocument(profileDocumentPath);
                memoPDF = new PdfDocument(memoDocumentPath);
                authPDF = new PdfDocument(authDocumentPath);

                document = new MergeDocument(profilePDF);
                document.Append(memoPDF, 1, 1);     // add only the first page
                document.Append(authPDF, 1, 1);     // add the SSN authorization document

            }


            if (wardDocArray.Length > 0 && ward.getStatus() != "X")    // if we have a letter of Guardianship
            {

                dh.getLetterOfGuardianship(wardDocArray, connection, ward.getWardNumber());                            // populate wardDocArray with the documents
                Console.WriteLine("Ward has {0} Letter(s) of Guardianship to copy", guardianLetterCount);

                //   B E G I N       M E R G E   O F   P D F    D O C U M E N T S           

                Console.WriteLine("output folder is: {0}", pdfTargetFolder);
                // loop over any additional documents here and merge with the Ward Profile pdf

                foreach (WardDocument w in wardDocArray)
                {
                    if (File.Exists(System.IO.Path.Combine(w.getStoragePath(), w.getDocPath())))
                    {
                        document.Append(System.IO.Path.Combine(w.getStoragePath(), w.getDocPath()));
                    }
                    else
                    {
                        System.IO.File.AppendAllText(@logFile, "Could not find ward document: " + w.getStoragePath() + "\\" + w.getDocPath() + " - " + DateTime.Now + "\r\n");
                    }
                }

            }
            else
            {
                Console.WriteLine("LOG: => No Letters of Guardianship found for ward {0}", ward.getWardNumber());
                System.IO.File.AppendAllText(@logFile, "Ward: " + ward.getWardNumber() + " Missing Letter of Gurardianship or CJIS Memorandum " + DateTime.Now + "\r\n");

            }

            ceTe.DynamicPDF.PageList pl = new PageList();
            pl = document.Pages;
            document2 = System.IO.Path.Combine(pdfTargetFolder, ward.getWardNumber() + "_CJISMemo_" + pl.Count.ToString() + ".pdf"); // name of the final output document
            document.Draw(document2);
            File.Delete(profileDocumentPath);      // delete the profile.pdf
            File.Delete(memoDocumentPath);         // delete the memo.pdf   
            File.Delete(authDocumentPath);         // delete the auth.pdf                            
            document2 = null;
            dh.updateCJISStatus(ward.getWardNumber(), guardianLetterCount, cjisMemoCount, connection,ward.getStatus());
            connection.Close();


        }
    }        /* end of class  */
}

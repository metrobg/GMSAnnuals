using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Oracle.DataAccess;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;
namespace fyiReporting.RdlCmd
{
    public class DocHelper
    {
        public Ward getWardInformation(int ward, OracleConnection connection)
        {
            Ward wardObj = null;

            string strSQL = " SELECT ward_number,employee_number,replace(replace(REPLACE(REPLACE(employee_name,',','_'),' ','_'),'\"',''),'*','') employee_name,ward_ward_status,";
            strSQL += "replace(replace(REPLACE(REPLACE(WARD_NAME,',','_'),' ','_'),'\"',''),'*','') WARD_NAME ,NVL(WARD_COURT_FILE_NO,'99-999') WARD_COURT_FILE_NO ";
            strSQL += "FROM ward w,employee e WHERE WARD_NUMBER =:p_ward and ";
            strSQL += "e.EMPLOYEE_NUMBER = w.WARD_RESPONSIBLE_EMPLOYEE";

            OracleCommand cmd = null;
            cmd = new OracleCommand();

            cmd.CommandText = strSQL;
            cmd.Connection = connection;
            OracleDataReader rs = null;

            try
            {
                OracleParameter p_ward = new OracleParameter();
               p_ward.Value = Convert.ToDecimal(ward);

                cmd.Parameters.Add(p_ward);

                cmd.ExecuteReader();
                rs = cmd.ExecuteReader();


                decimal ward_number;
                string ward_name;
                decimal employee_number;
                string employee_name;
                string file_number;
                string ward_status;

                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        ward_number     = rs.GetDecimal(0);
                        employee_number = rs.GetDecimal(1);
                        employee_name   = rs.GetString(2);
                        ward_name       = rs.GetString(4);
                        file_number     = rs.GetString(5);
                        ward_status     = rs.GetString(3);
                        
                        wardObj = new Ward(ward_number, ward_name, employee_number, employee_name,file_number,ward_status);

                    }
                }

            }
            catch (OracleException e)
            {
                Console.WriteLine("Exception Caught {0}", e.ToString());
                rs.Dispose();
                cmd.Dispose();
                Environment.Exit(-1);
            }
            rs.Dispose();
            cmd.Dispose();
            return wardObj;
        }




        public Array getWardBankDocs(WardDocument[] documentArray, OracleConnection connection, decimal ward, string enddate)
        {
            /*
             * 
 DOCUMENTS
	DOCNUM		1234
	DOCTYPE		OTHGL
	WARDNUM		4962
	VENDORNUM		0
	DOCID			07-2281
	DOCDATE		26-SEP-
	DATESCANNED		02-OCT-07	
	USERSCANNED		OPS$OPERATOR
	NOTES			null	
	MIMETYPE		application/pdf
	BASEPATH		0
	DOCPATH		2007\10\2\4962-OTHLGL-0--1234.pdf
               
*/
            string strSQL = "SELECT BASEPATH,DOCPATH,DOCNUM,DOCTYPE,WARDNUM,STORAGEPATH,replace(replace(REPLACE(REPLACE(WARD_NAME,',','_'),' ','_'),'\"',''),'*','') WARD_NAME ";
            strSQL += "FROM  DOCUMENTS d,  docpath dp,WARD W WHERE WARDNUM =:p_ward and ";
            strSQL += "DP.PATHNUM = D.BASEPATH  and W.WARD_NUMBER = D.WARDNUM ";
            strSQL += "AND (doctype = 'MELLON' OR doctype = 'BANK' or doctype = 'TRUST'  Or Doctype = 'OTHTST') ";
            strSQL += "AND trunc(docdate) =   (SELECT MAX(docdate) FROM documents WHERE (doctype = 'MELLON' OR doctype = 'BANK' or doctype = 'TRUST'  Or Doctype = 'OTHTST') ";
            strSQL += "AND extract(year from docdate) = extract(year from to_date(:p_enddate,'mm/dd/yyyy'))  AND wardnum = W.WARD_NUMBER)";
          
            OracleCommand cmd = null;
            OracleDataReader rs = null;

            try
            {
                cmd = new OracleCommand();
                cmd.CommandText = strSQL;
                cmd.Connection = connection;

                OracleParameter p_ward = new OracleParameter();
                p_ward.Value = Convert.ToDecimal(ward);
                cmd.Parameters.Add(p_ward);

                OracleParameter p_enddate = new OracleParameter();
                p_enddate.Value = enddate;
                cmd.Parameters.Add(p_enddate);

                cmd.ExecuteReader();
                rs = cmd.ExecuteReader();

                Console.WriteLine("Reading Ward Documents");
                WardDocument wardDoc;

                decimal basepath;
                string docpath;
                decimal docnum;
                string doctype;
                decimal wardnum;
                string storagepath;
                string wardName;

                int kount = 0;
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        basepath = rs.GetDecimal(0);
                        docpath = rs.GetString(1);
                        docnum = rs.GetDecimal(2);
                        doctype = rs.GetString(3);
                        wardnum = rs.GetDecimal(4);
                        storagepath = rs.GetString(5);
                        wardName = rs.GetString(6);

                        wardDoc = new WardDocument(basepath, docpath, docnum, doctype, wardnum, storagepath, wardName);
                        documentArray[kount] = wardDoc;
                        kount++;
                    }
                }

            }
            catch (OracleException e)
            {
                Console.WriteLine("Exception Caught {0}", e.ToString());
                rs.Dispose();
                cmd.Dispose();
                Environment.Exit(-1);
            }
            rs.Dispose();
            cmd.Dispose();
            return documentArray;
        }

        public Array getWardPhysicianReport(WardDocument[] documentArray, OracleConnection connection, decimal ward)
        {
            /*
             * 
 DOCUMENTS
	DOCNUM		1234
	DOCTYPE		OTHANN
	WARDNUM		4962
	VENDORNUM		0
	DOCID			07-2281
	DOCDATE		26-SEP-
	DATESCANNED		02-OCT-07	
	USERSCANNED		OPS$OPERATOR
	NOTES			null	
	MIMETYPE		application/pdf
	BASEPATH		0
	DOCPATH		2007\10\2\4962-OTHANN-0--1234.pdf
               
*/

            string strSQL = "SELECT BASEPATH,DOCPATH,DOCNUM,DOCTYPE,WARDNUM,STORAGEPATH,REPLACE(REPLACE(REPLACE(REPLACE(WARD_NAME,',','_'),' ','_'),'\"',''),'*','') WARD_NAME   ";
            strSQL += "FROM  DOCUMENTS d,  docpath dp,WARD W WHERE WARDNUM =:p_ward and DOCID  = '223' and DP.PATHNUM = D.BASEPATH  AND W.WARD_NUMBER = D.WARDNUM  ";
            strSQL += " And TRUNC(DOCDATE) > sysdate - (SELECT company_docltr_days FROM company)";
            

            OracleCommand cmd = null;
            OracleDataReader rs = null;

            try
            {
                cmd = new OracleCommand();
                cmd.CommandText = strSQL;
                cmd.Connection = connection;

                OracleParameter p_ward = new OracleParameter();
                p_ward.Value = Convert.ToDecimal(ward);
                cmd.Parameters.Add(p_ward);

                cmd.ExecuteReader();
                rs = cmd.ExecuteReader();

                Console.WriteLine("Reading Ward Physicians Report Documents");
                WardDocument wardDoc;

                decimal basepath;
                string docpath;
                decimal docnum;
                string doctype;
                decimal wardnum;
                string storagepath;
                string wardName;

                int kount = 0;
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        basepath = rs.GetDecimal(0);
                        docpath = rs.GetString(1);
                        docnum = rs.GetDecimal(2);
                        doctype = rs.GetString(3);
                        wardnum = rs.GetDecimal(4);
                        storagepath = rs.GetString(5);
                        wardName = rs.GetString(6);

                        wardDoc = new WardDocument(basepath, docpath, docnum, doctype, wardnum, storagepath, wardName);
                        documentArray[kount] = wardDoc;
                        kount++;
                    }
                }

            }
            catch (OracleException e)
            {
                Console.WriteLine("Exception Caught {0}", e.ToString());
                rs.Dispose();
                cmd.Dispose();
                Environment.Exit(-1);
            }
            rs.Dispose();
            cmd.Dispose();
            return   documentArray;
        }

 


        public Array getCJISMemo(WardDocument[] documentArray, OracleConnection connection, decimal ward)
        {
            /*
             * 
 DOCUMENTS
	DOCNUM		1234
	DOCTYPE		GRDLET
	WARDNUM		4962
	VENDORNUM		0
	DOCID			07-2281
	DOCDATE		26-SEP-
	DATESCANNED		02-OCT-07	
	USERSCANNED		OPS$OPERATOR
	NOTES			null	
	MIMETYPE		application/pdf
	BASEPATH		0
	DOCPATH		2007\10\2\4962-OTHANN-0--1234.pdf
               
*/

            string strSQL =   "select q.* from ( ";
                    strSQL += "SELECT docnum ,doctype ,WARDNUM,vendornum,DOCID,DOCDATE,DATESCANNED,USERSCANNED,NOTES,MIMETYPE ";
                    strSQL += ",BASEPATH,(select storagepath from docpath where pathnum = basepath) ||''||docpath docpath FROM DOCUMENTS ";
                    strSQL += "where DOCTYPE = 'CJIS' and wardnum = :p_ward  order by docdate desc) q where rownum = 1";


            OracleCommand cmd = null;
            OracleDataReader rs = null;

            try
            {
                cmd = new OracleCommand();
                cmd.CommandText = strSQL;
                cmd.Connection = connection;

                OracleParameter p_ward = new OracleParameter();
                p_ward.Value = Convert.ToDecimal(ward);
                cmd.Parameters.Add(p_ward);

                cmd.ExecuteReader();
                rs = cmd.ExecuteReader();

                Console.WriteLine("Looking for Letter of Guardianship");
                WardDocument wardDoc;

                decimal basepath;
                string docpath;
                decimal docnum;
                string doctype;
                decimal wardnum;
                string storagepath;
                string wardName;

                int kount = 0;
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        basepath = rs.GetDecimal(10);
                        docpath = rs.GetString(11);
                        docnum = rs.GetDecimal(0);
                        doctype = rs.GetString(1);
                        wardnum = rs.GetDecimal(2);
                        storagepath = rs.GetString(11);
                        wardName = "Ward Name";

                        wardDoc = new WardDocument(basepath, docpath, docnum, doctype, wardnum, storagepath, wardName);
                        documentArray[kount] = wardDoc;
                        kount++;
                    }
                }

            }
            catch (OracleException e)
            {
                Console.WriteLine("Exception Caught {0}", e.ToString());
                rs.Dispose();
                cmd.Dispose();
                Environment.Exit(-1);
            }
            rs.Dispose();
            cmd.Dispose();
            return documentArray;
        }
        public Array getLetterOfGuardianship(WardDocument[] documentArray, OracleConnection connection, decimal ward)
        {
            /*
             * 
 DOCUMENTS
	DOCNUM		1234
	DOCTYPE		GRDLET
	WARDNUM		4962
	VENDORNUM		0
	DOCID			07-2281
	DOCDATE		26-SEP-
	DATESCANNED		02-OCT-07	
	USERSCANNED		OPS$OPERATOR
	NOTES			null	
	MIMETYPE		application/pdf
	BASEPATH		0
	DOCPATH		2007\10\2\4962-OTHANN-0--1234.pdf
               
*/

            string strSQL =   "select q.* from ( ";
                    strSQL += "SELECT docnum ,doctype ,WARDNUM,vendornum,DOCID,DOCDATE,DATESCANNED,USERSCANNED,NOTES,MIMETYPE ";
                    strSQL += ",BASEPATH,(select storagepath from docpath where pathnum = basepath) ||''||docpath docpath FROM DOCUMENTS ";
                    strSQL += "where DOCTYPE = 'GRDLET' and docid = 24 and wardnum = :p_ward  order by docdate desc) q where rownum = 1";
              /*
                    strSQL += " union select q.* from ( ";
                    strSQL += "SELECT docnum ,doctype ,WARDNUM,vendornum,DOCID,DOCDATE,DATESCANNED,USERSCANNED,NOTES,MIMETYPE ";
                    strSQL += ",BASEPATH,(select storagepath from docpath where pathnum = basepath) ||''||docpath docpath FROM DOCUMENTS ";
                    strSQL += " where DOCTYPE = 'CJIS' and wardnum = :p_ward  order by docdate desc) q where rownum = 1 order by 2 asc";
               */

            OracleCommand cmd = null;
            OracleDataReader rs = null;

            try
            {
                cmd = new OracleCommand();
                cmd.CommandText = strSQL;
                cmd.Connection = connection;

                OracleParameter p_ward = new OracleParameter();
                p_ward.Value = Convert.ToDecimal(ward);
                cmd.Parameters.Add(p_ward);

                cmd.ExecuteReader();
                rs = cmd.ExecuteReader();

                Console.WriteLine("Looking for Letter of Guardianship");
                WardDocument wardDoc;

                decimal basepath;
                string docpath;
                decimal docnum;
                string doctype;
                decimal wardnum;
                string storagepath;
                string wardName;

                int kount = 0;
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        basepath = rs.GetDecimal(10);
                        docpath = rs.GetString(11);
                        docnum = rs.GetDecimal(0);
                        doctype = rs.GetString(1);
                        wardnum = rs.GetDecimal(2);
                        storagepath = rs.GetString(11);
                        wardName = "Ward Name";

                        wardDoc = new WardDocument(basepath, docpath, docnum, doctype, wardnum, storagepath, wardName);
                        documentArray[kount] = wardDoc;
                        kount++;
                    }
                }

            }
            catch (OracleException e)
            {
                Console.WriteLine("Exception Caught {0}", e.ToString());
                rs.Dispose();
                cmd.Dispose();
                Environment.Exit(-1);
            }
            rs.Dispose();
            cmd.Dispose();
            return documentArray;
        }


       

        public decimal getCurrentPath(OracleConnection connection)
        {
            string strSQL = "SELECT CURRENTPATHNUM FROM DOCCURRENTPATH";
            OracleCommand cmd = null;
            OracleDataReader rs = null;
            decimal pathnumber = -1;

            try
            {

                cmd = new OracleCommand();
                cmd.CommandText = strSQL;
                cmd.Connection = connection;
                cmd.ExecuteReader();

                rs = cmd.ExecuteReader();
                //Console.WriteLine("Reading CURRENTPATHNUMBER");

                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        pathnumber = rs.GetDecimal(0);
                    }
                }
            }
            catch (OracleException)
            {
                Console.WriteLine("Error reading the current path from the database");
                cmd.Dispose();
                rs.Dispose();
                Environment.Exit(-1);
            }

            cmd.Dispose();
            rs.Dispose();
            return pathnumber;
        }


      

       

      
        public String getCaseWorkerFolderName(int employee_number, OracleConnection connection)
        {
            
            string strSQL = " SELECT EMPLOYEE_NAME,EMPLOYEE_number,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(FLIP_NAME(EMPLOYEE_NAME),',','_'),' ','_'),'\"',''),'*',''),'/','_')||'_'||EMPLOYEE_number ";
            strSQL += "EMPLOYEE_name  FROM EMPLOYEE WHERE EMPLOYEE_number = " + employee_number;
            OracleCommand cmd = null;
            OracleDataReader rs = null;
            string pdfFolderName = null;

            try
            {
                cmd = new OracleCommand();
                cmd.CommandText = strSQL;
                cmd.Connection = connection;
                cmd.ExecuteReader();
                rs = cmd.ExecuteReader();

                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        pdfFolderName = rs.GetOracleString(0).ToString();
                    }
                }
            }
            catch (OracleException e)
            {
                Console.WriteLine("Exception Caught {0}", e.ToString());
                rs.Dispose();
                cmd.Dispose();
                Environment.Exit(-1);
            }
            rs.Dispose();
            cmd.Dispose();
            return pdfFolderName;
        }

        public int updatePlanStatus(decimal ward_number, OracleConnection connection)
        {
            
            string strSQL = "update annual_plan set SV_COMPLETED = 'X' where annual_plan_ward = " + ward_number;
             
            OracleCommand cmd = null;
            int rc = 0;
            try
            {
                cmd = new OracleCommand();
                cmd.CommandText = strSQL;
                cmd.Connection = connection;
                cmd.ExecuteReader();
                rc = cmd.ExecuteNonQuery();

                
               
            }
            catch (OracleException e)
            {
                Console.WriteLine("Exception Caught {0}", e.ToString());
               
                cmd.Dispose();
                Environment.Exit(-1);
            }
            
            return rc;
        }

        public int updateCJISStatus(decimal ward_number, decimal grdlet_count, decimal memo_count, OracleConnection connection)
        {

            string strSQL = "update WARD_CJIS set CJIS_MEMO_PRINTED = " + memo_count + " ,GRDLET_PRINTED = " + grdlet_count; ;  
            strSQL +=       " ,DATE_PRINTED = SYSDATE WHERE CJIS_WARD_NUMBER = " + ward_number;
            OracleCommand cmd = null;
            int rc = 0;
            try
            {
                cmd = new OracleCommand(strSQL, connection);
                rc  = cmd.ExecuteNonQuery();  
                 

            }
            catch (OracleException e)
            {
                Console.WriteLine("Exception Caught {0}", e.ToString());

                cmd.Dispose();
                Environment.Exit(-1);
            }

            return rc;
        }
        public int updateCJISStatus(decimal ward_number, decimal grdlet_count, decimal memo_count, OracleConnection connection,string status)
        {
            string strSQL = "update WARD_CJIS set CJIS_MEMO_PRINTED = " + memo_count + " ,GRDLET_PRINTED = " + grdlet_count;  
            if (status == "X")
            {
                strSQL += " ,DATE_REMOVED = SYSDATE WHERE CJIS_WARD_NUMBER = " + ward_number;
            }
            else
            {
                strSQL += " ,DATE_PRINTED = SYSDATE WHERE CJIS_WARD_NUMBER = " + ward_number;
            }
                
                OracleCommand cmd = null;

            int rc = 0;
            try
            {
                cmd = new OracleCommand(strSQL, connection);
                rc  = cmd.ExecuteNonQuery();


            }
            catch (OracleException e)
            {
                Console.WriteLine("Exception Caught {0}", e.ToString());

                cmd.Dispose();
                Environment.Exit(-1);
            }

            return rc;
        }
        public int updateAnnualAccounting(decimal ward_number, OracleConnection connection)
        {
            
            string strSQL = "insert into Annual_Accounting (billsel_ward,processed,date_processed) values(" + ward_number + ",'Y',SYSDATE)";

            OracleCommand cmd = null;
            int rc = 0;
            try
            {
                cmd = new OracleCommand();
                cmd.CommandText = strSQL;
                cmd.Connection = connection;
                rc = cmd.ExecuteNonQuery();

            }
            catch (OracleException e)
            {
                Console.WriteLine("Exception Caught {0}", e.ToString());

                cmd.Dispose();
                Environment.Exit(-1);
            }

            return rc;
        }
        public decimal getBankStatementCount(decimal ward, OracleConnection connection, string enddate)
        {
            string strSQL = "SELECT count(*) FROM DOCUMENTS D WHERE WARDNUM =:p_ward and ";
            strSQL += "(doctype = 'MELLON' or doctype ='BANK'  or doctype = 'TRUST' Or Doctype = 'OTHTST' )";
            strSQL += "AND trunc(docdate) =   (SELECT MAX(docdate) FROM documents WHERE (doctype = 'MELLON' OR doctype = 'BANK'  or doctype = 'TRUST'  Or Doctype = 'OTHTST') ";
            strSQL += "AND extract(year from docdate) = extract(year from to_date(:p_enddate,'mm/dd/yyyy'))  AND wardnum = D.WARDNUM)";

            OracleCommand cmd = null;
            cmd = new OracleCommand();

            cmd.CommandText = strSQL;
            cmd.Connection = connection;
            OracleDataReader rs = null;

                       
            decimal seq = 0;
            try
            {
                OracleParameter p_ward = new OracleParameter();
                p_ward.Value = Convert.ToDecimal(ward);
                cmd.Parameters.Add(p_ward);

                OracleParameter p_enddate = new OracleParameter();
                p_enddate.Value = enddate;
                cmd.Parameters.Add(p_enddate);

                cmd.ExecuteReader();
                rs = cmd.ExecuteReader();
              
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        seq = rs.GetDecimal(0);

                    }
                }
            }
            catch (OracleException e)
            {
                Console.WriteLine("Error reading Document Types {0}", e.ToString());
                rs.Dispose();
                cmd.Dispose();
                connection.Close();
                Environment.Exit(-1);
            }
            rs.Dispose();
            cmd.Dispose();
            return seq;


        }
        public decimal getPhysicianDocumentCount(decimal ward, OracleConnection connection)
        {
             
                    string strSQL = "SELECT COUNT(*) FROM documents d ";
                    strSQL += "WHERE DOCID = '223' AND WARDNUM = :p_ward ";
                    strSQL += "And TRUNC(DOCDATE)>sysdate - (SELECT company_docltr_days FROM company)";
                   

            OracleCommand cmd = null;
            cmd = new OracleCommand();
            cmd.CommandText = strSQL;
            cmd.Connection = connection;
            OracleDataReader rs = null;
            decimal seq = 0;
            try
            {
                OracleParameter p_ward = new OracleParameter();
                p_ward.Value = Convert.ToDecimal(ward);
                cmd.Parameters.Add(p_ward);

                cmd.ExecuteReader();
                rs = cmd.ExecuteReader();

                //Console.WriteLine("Count Document Types");
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        seq = rs.GetDecimal(0);

                    }
                }
            }
            catch (OracleException e)
            {
                Console.WriteLine("Error reading Document Types {0}", e.ToString());
                rs.Dispose();
                cmd.Dispose();
                connection.Close();
                Environment.Exit(-1);
            }
            rs.Dispose();
            cmd.Dispose();
            return seq;


        }

        public decimal getguardianLetterCount(decimal ward, OracleConnection connection)
        {

            string strSQL = "select count(*) from ( ";
            strSQL += "SELECT docnum ,doctype ,WARDNUM,vendornum,DOCID,DOCDATE,DATESCANNED,USERSCANNED,NOTES,MIMETYPE ";
            strSQL += ",BASEPATH,(select storagepath from docpath where pathnum = basepath) ||''||docpath docpath FROM DOCUMENTS ";
            strSQL += "where DOCTYPE = 'GRDLET' and docid = 24 and wardnum = :p_ward  order by docdate desc) q where rownum = 1";


            OracleCommand cmd = null;
            cmd = new OracleCommand();
            cmd.CommandText = strSQL;
            cmd.Connection = connection;
            OracleDataReader rs = null;
            decimal seq = 0;
            try
            {
                OracleParameter p_ward = new OracleParameter();
                p_ward.Value = Convert.ToDecimal(ward);
                cmd.Parameters.Add(p_ward);

                //cmd.ExecuteReader();
                rs = cmd.ExecuteReader();

                //Console.WriteLine("Count Document Types");
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        seq = rs.GetDecimal(0);

                    }
                }
            }
            catch (OracleException e)
            {
                Console.WriteLine("Error counting letters of guardianship {0}", e.ToString());
                rs.Dispose();
                cmd.Dispose();
                connection.Close();
                Environment.Exit(-1);
            }
            rs.Dispose();
            cmd.Dispose();
            return seq;


        }
        public decimal getCJISMemoCount(decimal ward, OracleConnection connection)
        {

            string strSQL = "select count(*) from ( ";
            strSQL += "SELECT docnum ,doctype ,WARDNUM,vendornum,DOCID,DOCDATE,DATESCANNED,USERSCANNED,NOTES,MIMETYPE ";
            strSQL += ",BASEPATH,(select storagepath from docpath where pathnum = basepath) ||''||docpath docpath FROM DOCUMENTS ";
            strSQL += "where DOCTYPE = 'CJIS' and wardnum = :p_ward  order by docdate desc) q where rownum = 1";


            OracleCommand cmd = null;
            cmd = new OracleCommand();
            cmd.CommandText = strSQL;
            cmd.Connection = connection;
            OracleDataReader rs = null;
            decimal seq = 0;
            try
            {
                OracleParameter p_ward = new OracleParameter();
                p_ward.Value = Convert.ToDecimal(ward);
                cmd.Parameters.Add(p_ward);

                //cmd.ExecuteReader();
                rs = cmd.ExecuteReader();

                //Console.WriteLine("Count Document Types");
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        seq = rs.GetDecimal(0);

                    }
                }
            }
            catch (OracleException e)
            {
                Console.WriteLine("Error counting CJIS Memo {0}", e.ToString());
                rs.Dispose();
                cmd.Dispose();
                connection.Close();
                Environment.Exit(-1);
            }
            rs.Dispose();
            cmd.Dispose();
            return seq;


        }
        public decimal getWardDocumentCount(int ward, OracleConnection connection)
        {
            string strSQL = "SELECT count(*) FROM DOCUMENTS WHERE WARDNUM = :p_ward and DOCPATH not like '%MULTIPLE%'";
            
            OracleCommand cmd = null;
            cmd = new OracleCommand();
            cmd.CommandText = strSQL;
            cmd.Connection = connection;

            OracleDataReader rs = null;
            decimal seq = 0;
            try
            {
                OracleParameter p_ward = new OracleParameter();
                p_ward.Value = Convert.ToDecimal(ward);
                cmd.Parameters.Add(p_ward);
               
                //cmd.ExecuteReader();
                rs = cmd.ExecuteReader();

                //Console.WriteLine("Count Document Types");
                if (rs.HasRows)
                {
                    while (rs.Read())
                    {
                        seq = rs.GetDecimal(0);

                    }
                }
            }
            catch (OracleException e)
            {
                Console.WriteLine("Error reading Document Types {0}", e.ToString());
                rs.Dispose();
                cmd.Dispose();
                connection.Close();
                Environment.Exit(-1);
            }
            rs.Dispose();
            cmd.Dispose();
            return seq;


        }

    }  // end of class
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Net;
using System.Xml;
//using Excel = Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Data.OleDb;
using Spire.Xls;
using Spire.Xls.Charts;


namespace EngieXML
{
    public class strSplit
    {
         public string resString;
        public strSplit() { }
    }
    class Program
    {
        private static XmlDocument CreateSoapEnvelope(int index, string strAzione, string strCodiceCliente, string strCodiceLottoAffido)
        {
            {   //Input Soap xml
                XmlDocument soapEnvelopeDocument = new XmlDocument();
                soapEnvelopeDocument.LoadXml(@"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:eic=""https://services.engie.it/ws/EICreditMgmtCM26.ws.provider:EAI_CM26"">" +
                   "<soapenv:Header/>" +
                   "<soapenv:Body>" +
                      "<eic:retrieveCreditPosition>" +
                         "<Input>" +
                            "<Codice_AdR>AXTR2505</Codice_AdR>" +
                            "<Azione>" + strAzione + "</Azione>" +
                            "<!--Optional:-->" +
                            "<Codice_Cliente>" + strCodiceCliente + "</Codice_Cliente>" +
                            "<!--Optional:-->" +
                            "<Codice_LottoAffido>" + strCodiceLottoAffido + "</Codice_LottoAffido>" +
                         "</Input>" +
                      "</eic:retrieveCreditPosition>" +
                   "</soapenv:Body>" +
                "</soapenv:Envelope>");
                return soapEnvelopeDocument;
            }
        }

        private static string GetDataDir()
        {
            var dataDir = System.Environment.CurrentDirectory;
            return dataDir;
        }

        private static HttpWebRequest CreateWebRequest(string url, string action)
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(url);
            webRequest.ContentType = "text/xml;charset=\"utf-8\"";
            webRequest.Accept = "text/xml";
            webRequest.Method = "POST";
            return webRequest;
        }

        private static void InsertSoapEnvelopeIntoWebRequest(XmlDocument soapEnvelopeXml, HttpWebRequest webRequest)
        {
            using (Stream stream = webRequest.GetRequestStream())
            {
                soapEnvelopeXml.Save(stream);
            }
        }

        public static string Selector(Spire.Xls.CellRange cell)
        {
            if (cell.Value2 == null)
                return "";
            if (cell.Value2.GetType().ToString() == "System.Double")
                return ((double)cell.Value2).ToString();
            else if (cell.Value2.GetType().ToString() == "System.String")
                return ((string)cell.Value2);
            else if (cell.Value2.GetType().ToString() == "System.Boolean")
                return ((bool)cell.Value2).ToString();
            else
                return "unknown";
        }

        public static string[] removeDuplicationValues(string[] values)
        {
            System.Collections.ArrayList result = new System.Collections.ArrayList();
            foreach (string s in values)
            {
                if (!result.Contains(s))
                {   if (s!="")
                    result.Add(s);
                }
            }
            return (string[])result.ToArray(typeof(string));
        }

        public static string[] SetStringArray(int intCount, string value)
        {
            string[] result = new String[intCount];
            for (int i = 0; i < intCount; i++)
            {
                result[i] = value;
            }
            return result;
        }

        static void Main(string[] args)
        {
            //WSDL Endpoint
            string _url = "https://some.com/wsdl"; 
            //WSDL action
            string _action = "EICreditMgmtCM26_ws_EAI_CM26_Port";
            // cleanXML - a variable that contains clean xml data, assists to form xml files
            string cleanXML = "";
            //startIndex, endIndex - variables that take part in handling Output Soap response files.
            int startIndex = 0;
            int endIndex = 0;
            // outputFile - a variable containing a name Output Soap response file
            string outputFile = "";
            try
            {
                List<XmlDocument> l = new List<XmlDocument>();
                string[] files = System.IO.Directory.GetFiles(GetDataDir(), "*.xlsx");

                if ((files.Length > 1))
                {

                    string logFile = GetDataDir() + "\\" + "logError" + "_" + DateTime.Now.ToString().Replace(":", "-").Replace("/", "-") + ".log";

                    if (!File.Exists(logFile))
                    {
                        StreamWriter txtFile = File.CreateText(logFile);
                        txtFile.Close();
                        txtFile = File.AppendText(logFile);
                        txtFile.WriteLine("There is more than 1 source file");
                        txtFile.Close();
                    }
                    System.Environment.Exit(1);
                }

                if (files.Length == 0)
                {

                    string logFile = GetDataDir() + "\\" + "logError" + "_" + DateTime.Now.ToString().Replace(":", "-").Replace("/", "-") + ".log";


                    if (!File.Exists(logFile))
                    {
                        StreamWriter txtFile = File.CreateText(logFile);
                        txtFile.Close();
                        txtFile = File.AppendText(logFile);
                        txtFile.WriteLine("Absent source Excel file");
                        txtFile.Close();
                    }
                    System.Environment.Exit(1);
                }

                string dir = GetDataDir();
                string[] xlFiles = Directory.GetFiles(dir, "*.xlsx");
                string inputFile = xlFiles[0];
                string strAzione;
                string strCodiceCliente;
                string strCodiceLottoAffido;
                int rowCount;

                Workbook xlWorkbook = new Workbook();
                xlWorkbook.LoadFromFile(inputFile);
                Worksheet sheet = xlWorkbook.Worksheets[0];


                Spire.Xls.CellRange xlRange = sheet.Range["A1:J1000"];
                rowCount = 1000;
                System.Array cellArray = xlRange.Cells.Cast<CellRange>().ToArray<CellRange>();
                List<CellRange> listOfCells = xlRange.Cells.Cast<CellRange>().ToList<CellRange>();
                string[] strArray = xlRange.Cells.Cast<CellRange>().Select(Selector).ToArray<string>();

                //xlWorkbook.Close(true);
                //xlApp.Quit();
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                GC.Collect();

                string[] arrAzione = SetStringArray(rowCount, "");
                string[] arrCodiceCliente = SetStringArray(rowCount, "");
                string[] arrCodiceLottoAffido = SetStringArray(rowCount, "");
                string[] arr_Azione_CodiceCliente_CodiceLottoAffido = SetStringArray(rowCount, "");
                // there is 3 columns in the input file
                for (var j = 1; ((j <= rowCount) && ((j * 3) < rowCount * 3)); j++)
                {
                    strAzione = strArray[(j * 3)];
                    arrAzione[j] = strAzione;
                    strCodiceCliente = strArray[(j * 3 + 1)];
                    arrCodiceCliente[j] = strCodiceCliente;
                    strCodiceLottoAffido = strArray[(j * 3 + 2)];
                    arrCodiceLottoAffido[j] = strCodiceLottoAffido;
                    arr_Azione_CodiceCliente_CodiceLottoAffido[j] = strAzione + "_" + strCodiceCliente + "_" + strCodiceLottoAffido;
                }

                //removing duplicates if they are in the input excel file
                string[] clean_arr_Azione_CodiceCliente_CodiceLottoAffido = SetStringArray(rowCount, "");
                clean_arr_Azione_CodiceCliente_CodiceLottoAffido = removeDuplicationValues(arr_Azione_CodiceCliente_CodiceLottoAffido);
                //-------------------------------------------------------
                //Assigning new rowCount after removing duplicates
                rowCount = clean_arr_Azione_CodiceCliente_CodiceLottoAffido.Length;
                //------------------------------------------------
                char[] dividers = { '_' };
                string[] ArrTemp = SetStringArray(3, "");
                string[] ArrNodividers = SetStringArray(rowCount, "");
                int k = 0;
                while (k < rowCount)
                {
                    ArrTemp = clean_arr_Azione_CodiceCliente_CodiceLottoAffido[k].Split(dividers);
                    strAzione = ArrTemp[0];
                    strCodiceCliente = ArrTemp[1];
                    strCodiceLottoAffido = ArrTemp[2];
                    XmlDocument doc = CreateSoapEnvelope(k, strAzione, strCodiceCliente, strCodiceLottoAffido);
                    l.Add(doc);
                    k++;
                }

                string[] cleanArrAzione = SetStringArray(rowCount, "");
                string[] cleanArrCodiceCliente = SetStringArray(rowCount, "");
                string[] cleanArrCodiceLottoAffido = SetStringArray(rowCount, "");
                string[] soapResult = SetStringArray(rowCount, "");

                XmlDocument[] soapEnvelopeXml = l.ToArray();

                //United output file for all xml files
                string UnitedOutputFile = GetDataDir() + "\\" + "united" + "_" + "output" + "_file" + ".xml";
                File.Delete(UnitedOutputFile);
                StreamWriter txtFile_united = File.CreateText(UnitedOutputFile);
                txtFile_united.WriteLine("<?xml version=" + "'1.0'" + "?><messaggi>");
                txtFile_united.Close();
                for (int j = 0; j < rowCount; ++j)
                {
                    HttpWebRequest webRequest = CreateWebRequest(_url, _action);
                    webRequest.Credentials = new System.Net.NetworkCredential("AXTR2505", "AXtr_01!");
                    InsertSoapEnvelopeIntoWebRequest(soapEnvelopeXml[j], webRequest);

                    // begin async call to web request.
                    IAsyncResult asyncResult = webRequest.BeginGetResponse(null, null);
                    asyncResult.AsyncWaitHandle.WaitOne();

                    // get the response from the completed web request.
                    using (WebResponse webResponse = webRequest.EndGetResponse(asyncResult))
                    {
                        using (StreamReader rd = new StreamReader(webResponse.GetResponseStream()))
                        {
                            soapResult[j] = rd.ReadToEnd();
                        }

                        startIndex = soapResult[j].IndexOf("<Body>");
                        endIndex = soapResult[j].IndexOf("</Body") - 6;
                        cleanXML = soapResult[j].Substring(startIndex + 6, endIndex - startIndex);
                        cleanXML = cleanXML.Replace("&lt;", "<").Replace("&gt;", ">");
                        outputFile = GetDataDir() + "\\" + "output" + "_" + clean_arr_Azione_CodiceCliente_CodiceLottoAffido[j].Substring(3) + ".xml";
                        txtFile_united.Close();
                        txtFile_united = File.AppendText(UnitedOutputFile);
                        txtFile_united.WriteLine(cleanXML);

                        if (!File.Exists(outputFile))
                        {
                            StreamWriter txtFile = File.CreateText(outputFile);
                            txtFile.Close();
                            txtFile = File.AppendText(outputFile);
                            txtFile.WriteLine(cleanXML);
                            txtFile.Close();
                        }
                        else
                        {
                            File.Delete(outputFile);
                            StreamWriter txtFile = File.CreateText(outputFile);
                            txtFile.Close();
                            txtFile = File.AppendText(outputFile);
                            txtFile.WriteLine(cleanXML);
                            txtFile.Close();
                        }
                    }
                }
                txtFile_united.WriteLine("</messaggi>");
                txtFile_united.Close();
            }
            
            catch (Exception ex)
            {
                // handle exception
                string logFile = GetDataDir() + "\\" + "logError" + "_" + DateTime.Now.ToString().Replace(":", "-").Replace("/", "-") + ".log";

                if (!File.Exists(logFile))
                {
                    StreamWriter txtFile = File.CreateText(logFile);
                    txtFile.Close();
                    txtFile = File.AppendText(logFile);
                    txtFile.WriteLine("There was an error: " + ex.Message  + "");
                    txtFile.Close();
                }
                System.Environment.Exit(1);
            }
        }
    }
}
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Xml;
using Spire.Xls;
using Spire.Xls.Collections;

namespace EngieXML
{
    // Logging Interface and Implementation
    public interface ILogger
    {
        void Log(string message);
    }

    public class FileLogger : ILogger
    {
        private readonly string _logFilePath;

        public FileLogger(string logFilePath)
        {
            _logFilePath = logFilePath;
        }

        public void Log(string message)
        {
            using (StreamWriter txtFile = File.AppendText(_logFilePath))
            {
                txtFile.WriteLine(message);
            }
        }
    }

    // SOAP Service Interface and Implementation
    public interface ISoapService
    {
        XmlDocument CreateSoapEnvelope(string action, string clientCode, string lotCode);
        string SendSoapRequest(XmlDocument soapEnvelope);
    }

    public class SoapService : ISoapService
    {
        private readonly string _url;
        private readonly string _action;
        private readonly NetworkCredential _credentials;

        public SoapService(string url, string action, NetworkCredential credentials)
        {
            _url = url;
            _action = action;
            _credentials = credentials;
        }

        public XmlDocument CreateSoapEnvelope(string action, string clientCode, string lotCode)
        {
            var soapEnvelopeDocument = new XmlDocument();
            soapEnvelopeDocument.LoadXml($@"<soapenv:Envelope xmlns:soapenv=""http://schemas.xmlsoap.org/soap/envelope/"" xmlns:eic=""https://services.engie.it/ws/EICreditMgmtCM26.ws.provider:EAI_CM26"">
                <soapenv:Header/>
                <soapenv:Body>
                    <eic:retrieveCreditPosition>
                        <Input>
                            <Codice_AdR>AXTR2505</Codice_AdR>
                            <Azione>{action}</Azione>
                            <Codice_Cliente>{clientCode}</Codice_Cliente>
                            <Codice_LottoAffido>{lotCode}</Codice_LottoAffido>
                        </Input>
                    </eic:retrieveCreditPosition>
                </soapenv:Body>
            </soapenv:Envelope>");
            return soapEnvelopeDocument;
        }

        public string SendSoapRequest(XmlDocument soapEnvelope)
        {
            HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(_url);
            webRequest.ContentType = "text/xml;charset=\"utf-8\"";
            webRequest.Accept = "text/xml";
            webRequest.Method = "POST";
            webRequest.Credentials = _credentials;

            using (Stream stream = webRequest.GetRequestStream())
            {
                soapEnvelope.Save(stream);
            }

            using (WebResponse webResponse = webRequest.GetResponse())
            {
                using (StreamReader rd = new StreamReader(webResponse.GetResponseStream()))
                {
                    return rd.ReadToEnd();
                }
            }
        }
    }

    // Excel File Processor
    public interface IExcelProcessor
    {
        string[] GetExcelData(string filePath);
    }

    public class ExcelProcessor : IExcelProcessor
    {
        public string[] GetExcelData(string filePath)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(filePath);
            Worksheet sheet = workbook.Worksheets[0];
            CellRange range = sheet.Range["A1:J1000"];
            return range.Cells.Cast<CellRange>().Select(cell => Selector(cell)).ToArray();
        }

        private static string Selector(CellRange cell)
        {
            if (cell.Value2 == null) return string.Empty;
            return cell.Value2 switch
            {
                double d => d.ToString(),
                string s => s,
                bool b => b.ToString(),
                _ => "unknown"
            };
        }
    }

    // File Manager for Output Files
    public interface IFileManager
    {
        void WriteOutputFile(string filePath, string content);
        void AppendToFile(string filePath, string content);
        void DeleteFile(string filePath);
    }

    public class FileManager : IFileManager
    {
        public void WriteOutputFile(string filePath, string content)
        {
            File.WriteAllText(filePath, content);
        }

        public void AppendToFile(string filePath, string content)
        {
            File.AppendAllText(filePath, content);
        }

        public void DeleteFile(string filePath)
        {
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
            }
        }
    }

    // Main Processor Class
    public class EngieProcessor
    {
        private readonly ILogger _logger;
        private readonly ISoapService _soapService;
        private readonly IExcelProcessor _excelProcessor;
        private readonly IFileManager _fileManager;

        public EngieProcessor(ILogger logger, ISoapService soapService, IExcelProcessor excelProcessor, IFileManager fileManager)
        {
            _logger = logger;
            _soapService = soapService;
            _excelProcessor = excelProcessor;
            _fileManager = fileManager;
        }

        public void Process(string excelFilePath, string unitedOutputFilePath)
        {
            try
            {
                string[] excelData = _excelProcessor.GetExcelData(excelFilePath);

                _fileManager.DeleteFile(unitedOutputFilePath);
                _fileManager.WriteOutputFile(unitedOutputFilePath, "<?xml version='1.0'?><messaggi>");

                for (int i = 0; i < excelData.Length; i += 3)
                {
                    string action = excelData[i];
                    string clientCode = excelData[i + 1];
                    string lotCode = excelData[i + 2];

                    XmlDocument soapEnvelope = _soapService.CreateSoapEnvelope(action, clientCode, lotCode);
                    string soapResponse = _soapService.SendSoapRequest(soapEnvelope);

                    string cleanXML = ExtractXmlBody(soapResponse);
                    string outputFile = $"{unitedOutputFilePath}_{clientCode}.xml";

                    _fileManager.WriteOutputFile(outputFile, cleanXML);
                    _fileManager.AppendToFile(unitedOutputFilePath, cleanXML);
                }

                _fileManager.AppendToFile(unitedOutputFilePath, "</messaggi>");
            }
            catch (Exception ex)
            {
                _logger.Log($"Error: {ex.Message}");
            }
        }

        private static string ExtractXmlBody(string xml)
        {
            int startIndex = xml.IndexOf("<Body>") + 6;
            int endIndex = xml.IndexOf("</Body>") - 6;
            string cleanXML = xml.Substring(startIndex, endIndex - startIndex);
            return cleanXML.Replace("&lt;", "<").Replace("&gt;", ">");
        }
    }

    // Main Program Entry Point
    class Program
    {
        static void Main(string[] args)
        {
            string excelFilePath = Path.Combine(Environment.CurrentDirectory, "input.xlsx");
            string logFilePath = Path.Combine(Environment.CurrentDirectory, $"logError_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}.log");
            string unitedOutputFilePath = Path.Combine(Environment.CurrentDirectory, "united_output_file.xml");

            ILogger logger = new FileLogger(logFilePath);
            ISoapService soapService = new SoapService("https://some.com/wsdl", "EICreditMgmtCM26_ws_EAI_CM26_Port", new NetworkCredential("AXTR2505", "AXtr_01!"));
            IExcelProcessor excelProcessor = new ExcelProcessor();
            IFileManager fileManager = new FileManager();

            EngieProcessor processor = new EngieProcessor(logger, soapService, excelProcessor, fileManager);
            processor.Process(excelFilePath, unitedOutputFilePath);
        }
    }
}

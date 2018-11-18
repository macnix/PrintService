using System;
using System.Collections.Generic;
using System.Linq;
using System.Management;
using System.Runtime.Serialization;
using System.ServiceModel;
using System.Text;

namespace wcfPrint
{
    // NOTE: You can use the "Rename" command on the "Refactor" menu to change the class name "PrintService" in code, svc and config file together.
    // NOTE: In order to launch WCF Test Client for testing this service, please select PrintService.svc or PrintService.svc.cs at the Solution Explorer and start debugging.
    public class PrintService : IPrintService
    {
        public List<String> GetPrinters()
        {
            try
            {
                var printers = new List<string>();
                ManagementScope objMS = new ManagementScope(ManagementPath.DefaultPath);
                objMS.Connect();
                SelectQuery objQuery = new SelectQuery("SELECT * FROM Win32_Printer");
                ManagementObjectSearcher objMOS = new ManagementObjectSearcher(objMS, objQuery);
                ManagementObjectCollection objMOC = objMOS.Get();
                string defaultPrinter = string.Empty;
                foreach (ManagementObject Printers in objMOC)
                {
                    // Default Printer.  
                    if (((bool?)Printers["Default"]) ?? false)
                    {
                        defaultPrinter = (Printers["Name"].ToString());
                    }
                    // Local and Network Printers.  
                    if (Convert.ToBoolean(Printers["Local"]) || Convert.ToBoolean(Printers["Network"]))
                    {
                        printers.Add(Printers["Name"].ToString());
                    }
                }
                return printers;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public void Print(string fileName, string printerName)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application { Visible = false };
                word.ActivePrinter = printerName;
                Microsoft.Office.Interop.Word.Document doc;
                doc = word.Documents.Open(fileName, ReadOnly: false, Visible: true);
                doc.PrintOut();
                doc.Close();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

    }
}

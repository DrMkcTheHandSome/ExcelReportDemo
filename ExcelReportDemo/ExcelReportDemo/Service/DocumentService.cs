using ClosedXML.Report;
using ExcelReportDemo.Interface;
using ExcelReportDemo.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelReportDemo.Service
{
    public class DocumentService : IDocumentService
    {
        public async Task<byte[]> ExportReportExcelFile(string fileServerPhysical)
        {
            return await Task.Run(() =>
            {
                var templateFullPath = Path.Combine(Directory.GetCurrentDirectory(), @"Templates", "sample_report.xlsx");
                var currentDirectory = new DirectoryInfo(Environment.CurrentDirectory);
                var cdnDirectory = new DirectoryInfo(
                    currentDirectory.FullName +
                    fileServerPhysical.Replace("{ProjectName}", currentDirectory.Name));

                // create if it doesnt exist
                Directory.CreateDirectory(cdnDirectory.FullName);

                fileServerPhysical = cdnDirectory.FullName;

                var outputFilePath = $"{fileServerPhysical}\\report-{Guid.NewGuid()}.xlsx";

                var template = new XLTemplate(templateFullPath);

                List<RiskViewModel> items = new List<RiskViewModel>()
            {
                new RiskViewModel()
                {
                    AgencyCode = "Bai1",
                    RiskStatusCode = "APP",
                    PolicyNumber = "123"
                },
                new RiskViewModel()
                {
                    AgencyCode = "Bai2",
                    RiskStatusCode = "ACT",
                    PolicyNumber = "123"
                },
                new RiskViewModel()
                {
                    AgencyCode = "Bai3",
                    RiskStatusCode = "APP",
                    PolicyNumber = "123"
                },
            };
                template.AddVariable("items", items);
                template.AddVariable("Company", "BlastAsia");
                template.AddVariable("Addr1", "The Orient Square, Emerald Ave,");
                template.AddVariable("Addr2", "San Antonio, Pasig, Metro Manila");
                template.AddVariable("City", "Jacksonville");
                template.AddVariable("Country", "USA");
                template.AddVariable("State", "Florida");
                template.AddVariable("Zip", "904");
                template.Generate();

                template.SaveAs(outputFilePath);
                byte[] packetFileBytes = File.ReadAllBytes(outputFilePath);
                if (Directory.Exists(cdnDirectory.FullName)) // Delete Generated Test Files
                {
                    Directory.Delete(cdnDirectory.FullName,true);
                }
                return packetFileBytes;
            });
        }
    }
}

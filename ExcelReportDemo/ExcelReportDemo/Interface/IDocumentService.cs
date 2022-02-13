using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelReportDemo.Interface
{
    public interface IDocumentService
    {
        Task<byte[]> ExportReportExcelFile(string fileServerPhysical);
    }
}

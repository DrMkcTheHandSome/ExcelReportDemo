using ClosedXML.Report;
using ExcelReportDemo.Interface;
using ExcelReportDemo.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelReportDemo.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ReportController : ControllerBase
    {
        private readonly ILogger<ReportController> _logger;
        private readonly IConfiguration _configuration;
        private readonly IDocumentService _documentService;

        public ReportController(
            IDocumentService documentService,
            ILogger<ReportController> logger,
            IConfiguration configuration)
        {
            _logger = logger;
            _configuration = configuration ?? throw new ArgumentNullException(nameof(configuration));
            _documentService = documentService ?? throw new ArgumentNullException(nameof(documentService));

        }


        [HttpGet("excelreport")]
        public async Task<ActionResult> ExportExcelReport()
        {
            try
            {
                var packetFile = await _documentService.ExportReportExcelFile(_configuration["FileServerPhysical"]);

                var packetReportFile = new FileResultFromStream(
                    $"excel_report.xlsx",
                    new MemoryStream(packetFile),
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                );

                return packetReportFile;
            }
            catch (Exception e)
            {
                return StatusCode(500, e.Message);
            }
        }
    }
}

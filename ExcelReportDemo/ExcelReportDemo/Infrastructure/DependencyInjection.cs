using ExcelReportDemo.Interface;
using ExcelReportDemo.Service;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RivTech.Centauri.Web.Infrastructure
{
    public static class DependencyInjection
    {
        public static IServiceCollection AddApplication(this IServiceCollection services)
        {
            services.AddScoped<IDocumentService, DocumentService>();
            return services;
        }
    }
}

using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ReadExcel
{
    class Dal
    {
        static IConfiguration _iconfiguration;
        public static void GetAppSettingsFile()
        {
            var builder = new ConfigurationBuilder()
                                 .SetBasePath(Directory.GetCurrentDirectory())
                                 .AddJsonFile("config.json", optional: false, reloadOnChange: true);
            _iconfiguration = builder.Build();
            _iconfiguration.GetConnectionString("SQLServerConnection");
            _iconfiguration.GetConnectionString("MySQLConnection");
        }
    }
}

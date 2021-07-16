using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CrawlerFipe
{
    class Program
    {
        static void Main(string[] args)
        {
            var consultaWeb= new ConsultaWeb();
            consultaWeb.ConsultaVeiculos();
            consultaWeb.Dispose();
        }
    }
}

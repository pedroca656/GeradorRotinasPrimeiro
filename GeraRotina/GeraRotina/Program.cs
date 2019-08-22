using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using GemBox.Spreadsheet;

namespace GeraRotina
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Iniciando aplicação...");
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            var workBook = new ExcelFile();
            ExcelWorksheet worksheet = workBook.Worksheets.Add("rotina1");

            worksheet.Cells[0, 0].Value = "Hora";
            worksheet.Cells[0, 1].Value = "Cômodo"; //C = Cozinha, Q = Quarto, S = Sala, B = Banheiro, F = Fora de Casa
            worksheet.Cells[0, 2].Value = "Dia Útil?";
            worksheet.Cells[0, 3].Value = "";

            var row = 1;

            for (int i = 0; i <= 10; i++)
            {
                
            }
        }
    }
}

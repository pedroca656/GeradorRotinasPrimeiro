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
                worksheet.Cells[1, 0].Value = new DateTime(1, 1, 1, 7, 15, 0).TimeOfDay;
                worksheet.Cells[1, 1].Value = "B";
                worksheet.Cells[1, 2].Value = true;

                worksheet.Cells[2, 0].Value = new DateTime(1, 1, 1, 8, 0, 0).TimeOfDay;
                worksheet.Cells[2, 1].Value = "Q";
                worksheet.Cells[2, 2].Value = true;

                worksheet.Cells[3, 0].Value = new DateTime(1, 1, 1, 8, 30, 0).TimeOfDay;
                worksheet.Cells[3, 1].Value = "C";
                worksheet.Cells[3, 2].Value = true;

                worksheet.Cells[4, 0].Value = new DateTime(1, 1, 1, 9, 0, 0).TimeOfDay;
                worksheet.Cells[4, 1].Value = "F";
                worksheet.Cells[4, 2].Value = true;

                worksheet.Cells[5, 0].Value = new DateTime(1, 1, 1, 12, 0, 0).TimeOfDay;
                worksheet.Cells[5, 1].Value = "S";
                worksheet.Cells[5, 2].Value = true;

                worksheet.Cells[6, 0].Value = new DateTime(1, 1, 1, 13, 0, 0).TimeOfDay;
                worksheet.Cells[6, 1].Value = "C";
                worksheet.Cells[6, 2].Value = true;

                worksheet.Cells[7, 0].Value = new DateTime(1, 1, 1, 13, 30, 0).TimeOfDay;
                worksheet.Cells[7, 1].Value = "F";
                worksheet.Cells[7, 2].Value = true;

                worksheet.Cells[8, 0].Value = new DateTime(1, 1, 1, 17, 30, 0).TimeOfDay;
                worksheet.Cells[8, 1].Value = "S";
                worksheet.Cells[8, 2].Value = true;

                worksheet.Cells[9, 0].Value = new DateTime(1, 1, 1, 19, 0, 0).TimeOfDay;
                worksheet.Cells[9, 1].Value = "B";
                worksheet.Cells[9, 2].Value = true;

                worksheet.Cells[10, 0].Value = new DateTime(1, 1, 1, 19, 30, 0).TimeOfDay;
                worksheet.Cells[10, 1].Value = "Q";
                worksheet.Cells[10, 2].Value = true;

                worksheet.Cells[11, 0].Value = new DateTime(1, 1, 1, 20, 0, 0).TimeOfDay;
                worksheet.Cells[11, 1].Value = "C";
                worksheet.Cells[11, 2].Value = true;

                worksheet.Cells[12, 0].Value = new DateTime(1, 1, 1, 22, 0, 0).TimeOfDay;
                worksheet.Cells[12, 1].Value = "S";
                worksheet.Cells[12, 2].Value = true;

                worksheet.Cells[1, 0].Value = new DateTime(1, 1, 1, 23, 30, 0).TimeOfDay;
                worksheet.Cells[1, 1].Value = "Q";
                worksheet.Cells[1, 2].Value = true;
            }
        }
    }
}

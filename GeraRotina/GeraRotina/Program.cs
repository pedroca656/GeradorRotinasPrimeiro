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
            if(args.Length < 2)
            {
                Console.WriteLine("Por favor, passe como parâmetros a quantidade de dias da rotina e o nome do arquivo final!");
                return;
            }

            var dias = 0;

            if (!int.TryParse(args[0], out dias))
            {
                Console.WriteLine("Por favor, informe números inteiros como dias!");
                return;
            }

            Console.WriteLine("Iniciando aplicação...");
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

            var workBook = new ExcelFile();
            ExcelWorksheet worksheet = workBook.Worksheets.Add("rotina1");

            worksheet.Cells[0, 0].Value = "Hora";
            worksheet.Cells[0, 1].Value = "Cômodo"; //C = Cozinha, Q = Quarto, S = Sala, B = Banheiro, F = Fora de Casa
            worksheet.Cells[0, 2].Value = "Dia Útil?";
            worksheet.Cells[0, 3].Value = "";

            var row = 1;

            for (int i = 0; i < dias; i++)
            {
                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 7, 15, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "B";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 8, 0, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "Q";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 8, 30, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "C";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 9, 0, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "F";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 12, 0, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "S";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 13, 0, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "C";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 13, 30, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "F";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 17, 30, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "S";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 19, 0, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "B";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 19, 30, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "Q";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 20, 0, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "C";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 22, 0, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "S";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                worksheet.Cells[row, 0].Value = new DateTime(1, 1, 1, 23, 30, 0).ToShortTimeString();
                worksheet.Cells[row, 1].Value = "Q";
                worksheet.Cells[row, 2].Value = "Sim";
                row++;

                Console.WriteLine("Dia " + i + " simulado...");
            }

            workBook.Save(args[1]+".xlsx");

            Console.WriteLine("Arquivo gerado!");
        }
    }
}

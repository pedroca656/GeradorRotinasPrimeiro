using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using NPOI.XSSF.UserModel;
using System.IO;

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

            var workBook = new XSSFWorkbook();
            var worksheet = workBook.CreateSheet("rotina1");


            var r = worksheet.CreateRow(0);
            r.CreateCell(1).SetCellValue("Dia Mês");
            r.CreateCell(1).SetCellValue("Hora");
            r.CreateCell(2).SetCellValue("Cômodo"); //C = Cozinha, Q = Quarto, S = Sala, B = Banheiro, F = Fora de Casa
            r.CreateCell(3).SetCellValue("Dia Útil?");

            var row = 1;

            for (int i = 0; i <= dias; i++)
            {
                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 7, 15, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("B");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 8, 0, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("Q");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 8, 30, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("C");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 9, 0, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("F");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 12, 0, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("S");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 13, 0, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("C");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 13, 30, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("F");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 17, 30, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("S");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 19, 0, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("B");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 19, 30, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("Q");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 20, 0, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("C");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 22, 0, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("S");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(new DateTime(2019, 1, 1).AddDays(i).Day.ToString());
                r.CreateCell(1).SetCellValue(new DateTime(1, 1, 1, 23, 30, 0).ToShortTimeString());
                r.CreateCell(2).SetCellValue("Q");
                r.CreateCell(3).SetCellValue("Sim");
                row++;

                Console.WriteLine("Dia " + i + " simulado...");
            }

            using (var fs = new FileStream(args[1] + ".xls", FileMode.Create, FileAccess.Write))
            {
                workBook.Write(fs);
            }

            Console.WriteLine("Arquivo gerado!");
        }
    }
}

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
        //teste
        static void Main(string[] args)
        {
            if(args.Length < 3)
            {
                Console.WriteLine("Por favor, passe como parâmetros a quantidade de dias da rotina e o nome do arquivo final, além do tipo de algoritmo!");
                return;
            }

            var dias = 0;

            if (!int.TryParse(args[0], out dias))
            {
                Console.WriteLine("Por favor, informe números inteiros como dias!");
                return;
            }

            var isDBSCAN = false;

            if (!bool.TryParse(args[2], out isDBSCAN))
            {
                Console.WriteLine("Por favor, informe se o tipo será dbscan!");
                return;
            }

            if (isDBSCAN)
            {
                executaParaDBSCAN(dias, args[1]);
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

        public static void executaParaDBSCAN(int dias, string nomeArquivo)
        {
            Console.WriteLine("Iniciando aplicação...");

            var workBook = new XSSFWorkbook();
            var worksheet = workBook.CreateSheet("rotina");


            var r = worksheet.CreateRow(0);
            //hora será em números inteiros
            r.CreateCell(0).SetCellValue("Hora");
            //comodo trocado para B = 100, Q = 200, C = 300, F = 400, S = 500 para poder executar o algoritmo
            r.CreateCell(1).SetCellValue("DiaSemana-Comodo"); //C = Cozinha, Q = Quarto, S = Sala, B = Banheiro, F = Fora de Casa
            //r.CreateCell(2).SetCellValue("Dia Mês");

            var row = 1;

            for (int i = 0; i <= dias; i++)
            {
                var diaSemana = getDiaSemana(new DateTime(2019, 1, 1).AddDays(i).DayOfWeek);                

                if (diaSemana == 4) //rotina somente de quarta, simulando um dia em que o indivíduo chega mais tarde em casa
                {
                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(07.15 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(100 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(08.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(200 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(08.50 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(300 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(09.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(400 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(12.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(500 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(13.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(300 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(13.50 + GetRandomDouble(-0.125, 0.125), 2));
                    r.CreateCell(1).SetCellValue(400 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(22.50 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(500 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(23.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(100 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(23.50 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(200 + (diaSemana * 1000));

                    row++;

                    Console.WriteLine("Dia " + i + " simulado...");
                    continue;
                }
                else if (diaSemana == 1) //rotina somente de domingo, simulando um almoço em família
                {
                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(09.15 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(100 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(09.75 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(200 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(10.25 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(300 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(12.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(500 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(14.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(300 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(16.00 + GetRandomDouble(-0.125, 0.125), 2));
                    r.CreateCell(1).SetCellValue(500 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(19.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(200 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(19.50 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(100 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(20.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(500 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(23.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(200 + (diaSemana * 1000));

                    row++;

                    Console.WriteLine("Dia " + i + " simulado...");
                    continue;
                }
                else if (diaSemana == 7) //rotina somente de sábado, simulando um dia mais "preguiçoso"
                {
                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(10.50 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(100 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(11.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(300 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(12.50 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(500 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(13.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(300 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(14.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(500 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(18.00 + GetRandomDouble(-0.125, 0.125), 2));
                    r.CreateCell(1).SetCellValue(100 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(19.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(300 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(20.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(500 + (diaSemana * 1000));

                    row++;

                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(System.Math.Round(01.00 + GetRandomDouble(-0.25, 0.25), 2));
                    r.CreateCell(1).SetCellValue(200 + (diaSemana * 1000));

                    row++;

                    Console.WriteLine("Dia " + i + " simulado...");
                    continue;
                }

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(07.15 + GetRandomDouble(-0.25, 0.25), 2));
                r.CreateCell(1).SetCellValue(100 + (diaSemana * 1000));

                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(08.00 + GetRandomDouble(-0.25, 0.25), 2));
                r.CreateCell(1).SetCellValue(200 + (diaSemana * 1000));
                
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(08.50 + GetRandomDouble(-0.25, 0.25), 2));
                r.CreateCell(1).SetCellValue(300 + (diaSemana * 1000));
                
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(09.00 + GetRandomDouble(-0.25, 0.25), 2));
                r.CreateCell(1).SetCellValue(400 + (diaSemana * 1000));
                
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(12.00 + GetRandomDouble(-0.25, 0.25), 2));
                r.CreateCell(1).SetCellValue(500 + (diaSemana * 1000));
                
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(13.00 + GetRandomDouble(-0.25, 0.25), 2));
                r.CreateCell(1).SetCellValue(300 + (diaSemana * 1000));
                
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(13.50 + GetRandomDouble(-0.125, 0.125), 2));
                r.CreateCell(1).SetCellValue(400 + (diaSemana * 1000));
                
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(17.50 + GetRandomDouble(-0.25, 0.25), 2));
                r.CreateCell(1).SetCellValue(500 + (diaSemana * 1000));
                
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(19.00 + GetRandomDouble(-0.25, 0.25), 2));
                r.CreateCell(1).SetCellValue(100 + (diaSemana * 1000));
                
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(19.50 + GetRandomDouble(-0.125, 0.125), 2));
                r.CreateCell(1).SetCellValue(200 + (diaSemana * 1000));
                
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(20.00 + GetRandomDouble(-0.25, 0.25), 2));
                r.CreateCell(1).SetCellValue(300 + (diaSemana * 1000));
                
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(22.00 + GetRandomDouble(-0.25, 0.25), 2));
                r.CreateCell(1).SetCellValue(500 + (diaSemana * 1000));
                
                row++;

                r = worksheet.CreateRow(row);
                r.CreateCell(0).SetCellValue(System.Math.Round(23.50 + GetRandomDouble(-0.25, 0.25), 2));
                r.CreateCell(1).SetCellValue(200 + (diaSemana * 1000));
                
                row++;

                Console.WriteLine("Dia " + i + " simulado...");
            }

            using (var fs = new FileStream(nomeArquivo + ".xls", FileMode.Create, FileAccess.Write))
            {
                workBook.Write(fs);
            }

            Console.WriteLine("Arquivo gerado!");
        }

        public static double GetRandomDouble(double minimum, double maximum)
        {
            Random random = new Random();
            //multiplica pelo maximo - minimo e soma o maximo para ter entre a range informada
            var result = random.NextDouble() * (maximum - minimum) + minimum;
            return result;
        }

        public static int getDiaSemana(DayOfWeek dia)
        {
            switch (dia)
            {
                case DayOfWeek.Sunday:
                    return 1;

                case DayOfWeek.Monday:
                    return 2;

                case DayOfWeek.Tuesday:
                    return 3;

                case DayOfWeek.Wednesday:
                    return 4;

                case DayOfWeek.Thursday:
                    return 5;

                case DayOfWeek.Friday:
                    return 6;

                case DayOfWeek.Saturday:
                    return 7;

                default:
                    return 0;
            }
        }
    }
}

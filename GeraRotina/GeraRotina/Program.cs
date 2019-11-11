using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Globalization;
using NPOI.XSSF.UserModel;
using System.IO;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;

namespace GeraRotina
{
    class Program
    {
        private static Dictionary<string, string> mapComodo = new Dictionary<string, string>()
        {
            {"B", "100"},
            {"Q", "200"},
            {"C", "300"},
            {"F", "400"},
            {"S", "500"},
        };


        static void Main(string[] args)
        {
            CultureInfo.CurrentCulture = CultureInfo.InvariantCulture;
            CultureInfo.CurrentUICulture = CultureInfo.InvariantCulture;

            if (args.Length < 3)
            {
                Console.WriteLine("Por favor, passe como parâmetros a quantidade de dias da rotina e o nome do arquivo final, além do tipo de algoritmo!");
                return;
            }

            var dias = 0;
            var isDBSCAN = false;
            var isLOF = false;
            //var isOPTICS = false;
            var percAnomalias = 0m;

            if (!int.TryParse(args[0], out dias))
            {
                Console.WriteLine("Por favor, informe números inteiros como dias!");
                return;
            }

            int algoritmo;
            var algoritmoNome = "";
            if (!int.TryParse(args[2], out algoritmo))
            {
                Console.WriteLine("Por favor, informe se o tipo será dbscan!");
                return;
            }

            isDBSCAN = algoritmo == 1;
            isLOF = algoritmo == 2;
            //isOPTICS = algoritmo == 3;

            if (args.Length >= 4 && !decimal.TryParse(args[3], out percAnomalias))
            {
                throw new ArgumentException("Parâmetro inválido", nameof(percAnomalias));
            }



            //if (isDBSCAN)
            //{
            //    executaParaDBSCAN(dias, args[1]);
            //    return;
            //}


            Console.WriteLine("Iniciando aplicação...");

            var workBook = new XSSFWorkbook();
            var worksheet = workBook.CreateSheet("rotina1");

            // Cabeçalho
            var r = worksheet.CreateRow(0);
            switch (algoritmo)
            {
                case 1: // DBSCAN
                case 2: // LOF
                    //hora será em números inteiros
                    r.CreateCell(0).SetCellValue("Hora");
                    //comodo trocado para B = 100, Q = 200, C = 300, F = 400, S = 500 para poder executar o algoritmo
                    r.CreateCell(1).SetCellValue("DiaSemana-Comodo"); //C = Cozinha, Q = Quarto, S = Sala, B = Banheiro, F = Fora de Casa
                    //r.CreateCell(2).SetCellValue("Dia Mês");
                    break;

                //case 3: // OPTICS

                default:
                    r.CreateCell(0).SetCellValue("Dia Mês");
                    r.CreateCell(1).SetCellValue("Hora");
                    r.CreateCell(2).SetCellValue("Cômodo"); //C = Cozinha, Q = Quarto, S = Sala, B = Banheiro, F = Fora de Casa
                    r.CreateCell(3).SetCellValue("Dia Útil?");

                    break;
            }

            var row = 1;
            var dataBase = new DateTime(2019, 1, 1);
            for (var i = 0; i <= dias; i++)
            {
                var rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(7, 15, 0), "B", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(8, 0, 0), "Q", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(8, 30, 0), "C", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(9, 0, 0), "F", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(12, 0, 0), "S", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(13, 0, 0), "C", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(13, 30, 0), "F", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(17, 30, 0), "S", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(19, 0, 0), "B", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(19, 30, 0), "Q", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(20, 0, 0), "C", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(22, 0, 0), "S", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                rotina = new Rotina(dataBase.AddDays(i), new TimeSpan(23, 30, 0), "Q", true);
                WriteLine(worksheet, row++, rotina, algoritmo);

                Console.WriteLine("Dia " + i + " simulado...");
            }

            var dir = Path.Combine(AppContext.BaseDirectory, "output");
            Directory.CreateDirectory(dir);

            switch (algoritmo)
            {
                case 1:
                    algoritmoNome = "DBSCAN";
                    break;
                case 2:
                    algoritmoNome = "LOF";
                    break;
                case 3:
                    algoritmoNome = "OPTICS";
                    break;
                default:
                    algoritmoNome = "";
                    break;
            }
            var fileName = Path.Combine(dir, args[1] + algoritmoNome + ".xlsx");

            //using (var fs = new FileStream(args[1] + ".xls", FileMode.Create, FileAccess.Write))
            using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
            {
                workBook.Write(fs);
            }

            Console.WriteLine($"Arquivo gerado! {fileName}");

#if DEBUG
            // Abrir arquivo
            System.Diagnostics.Process.Start(fileName);
#endif
        }

        private static void executaParaDBSCAN(int dias, string nomeArquivo)
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

        private static double GetRandomDouble(double minimum, double maximum)
        {
            Random random = new Random();
            //multiplica pelo maximo - minimo e soma o maximo para ter entre a range informada
            var result = random.NextDouble() * (maximum - minimum) + minimum;
            return result;
        }

        private static int getDiaSemana(DayOfWeek dia)
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


        private static void WriteLine(ISheet worksheet, int row, Rotina rotina, int algoritmo)
        {
            var r = worksheet.CreateRow(row);

            switch (algoritmo)
            {
                case 1: // DBSCAN
                case 2: // LOF
                    r = worksheet.CreateRow(row);
                    r.CreateCell(0).SetCellValue(rotina.DiaMes.Add(rotina.Hora).ToShortTimeString());
                    r.CreateCell(1).SetCellValue(getDiaSemana(rotina.DiaMes.DayOfWeek) * 1000 + int.Parse(mapComodo[rotina.Comodo]));

                    break;

                //case 3: // OPTICS

                default:
                    r.CreateCell(0).SetCellValue(rotina.DiaMes.ToShortDateString());
                    r.CreateCell(1).SetCellValue(rotina.DiaMes.Add(rotina.Hora).ToShortTimeString());
                    r.CreateCell(2).SetCellValue(rotina.Comodo);
                    r.CreateCell(3).SetCellValue(rotina.isDiaUtil ? "Sim" : "Não");
                    break;
            }


        }

        private class Rotina
        {
            public DateTime DiaMes { get; set; }
            public TimeSpan Hora { get; set; }
            public string Comodo { get; set; }
            public bool isDiaUtil { get; set; }

            public Rotina(DateTime diaMes, TimeSpan hora, string comodo, bool isDiaUtil)
            {
                DiaMes = diaMes;
                Hora = hora;
                Comodo = comodo;
                this.isDiaUtil = isDiaUtil;
            }
        }
    }
}

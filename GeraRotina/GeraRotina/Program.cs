using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;

namespace GeraRotina
{
    class Program
    {
        private static string separator = ",";

        private static Dictionary<string, string> mapComodo = new Dictionary<string, string>()
        {
            {"B", "100"},
            {"Q", "200"},
            {"C", "300"},
            {"F", "400"},
            {"S", "500"},
        };

        private static Random random = new Random();


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

            Console.WriteLine("Iniciando aplicação...");

            var sb = new List<string>();

            // Cabeçalho
            var s = new List<string>();
            switch (algoritmo)
            {
                case 1: // DBSCAN
                case 2: // LOF
                    //hora será em números inteiros
                    s.Add("Hora");
                    //comodo trocado para B = 100, Q = 200, C = 300, F = 400, S = 500 para poder executar o algoritmo
                    s.Add("DiaSemana-Comodo");
                    //r.CreateCell(2).SetCellValue("Dia Mês");
                    break;

                //case 3: // OPTICS

                default:
                    s.Add("Dia Mês");
                    s.Add("Hora");
                    s.Add("Cômodo");
                    s.Add("Dia Útil?");

                    break;
            }
            sb.Add(string.Join(separator, s));

            var rndMin = -15;
            var rndMax = 15;

            //Rotina1(dias, countAnomalias, rndMin, rndMax, sb, algoritmo);
            var anomaliasGeradas = Rotina2(dias, percAnomalias, rndMin, rndMax, sb, algoritmo);

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

            if (System.Environment.MachineName.Contains("RODRIGO"))
            {
                dir = Path.Combine(@"C:\Users\rodri\Google Drive\Datasets-beta");
            }


            var fileName = Path.Combine(dir, args[1] + algoritmoNome + $"__{anomaliasGeradas}__" + ".csv");
            using (var fs = new StreamWriter(fileName))
            {
                fs.Write(string.Join("\n", sb));
            }

            Console.WriteLine($"Arquivo gerado! {fileName}");

#if DEBUG
            // Abrir arquivo
            System.Diagnostics.Process.Start(fileName);
#endif
        }

        private static int Rotina2(int dias, decimal percAnomalias, int rndMin, int rndMax, List<string> sb, int algoritmo)
        {
            var rotinas = new List<Rotina>();
            var dataBase = new DateTime(2019, 1, 1);
            var time = new TimeSpan(0, 0, 0);

            for (int i = 0; i <= dias; i++)
            {
                var diaSemana = getDiaSemana(dataBase.DayOfWeek);
                if (diaSemana == 4) //rotina somente de quarta, simulando um dia em que o indivíduo chega mais tarde em casa
                {
                    time = getTimeSpan(time, new TimeSpan(7, 10, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", true));

                    time = getTimeSpan(time, new TimeSpan(08, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(08, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(09, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "F", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(12, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(13, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    //time = getTimeSpan(time, new TimeSpan(13.50 + GetRandomDouble(-0.125, 0.125), 2));
                    time = getTimeSpan(time, new TimeSpan(13, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "F", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(22, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(23, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(23, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    Console.WriteLine("Dia " + i + " simulado...");

                    continue;
                }
                else if (diaSemana == 1) //rotina somente de domingo, simulando um almoço em família
                {
                    time = getTimeSpan(time, new TimeSpan(09, 15, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(09, 50, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(10, 15, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(12, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(14, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    //time = getTimeSpan(time, new TimeSpan(16.00 + GetRandomDouble(-0.125, 0.125), 2));
                    time = getTimeSpan(time, new TimeSpan(16, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(19, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(19, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(20, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(23, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    Console.WriteLine("Dia " + i + " simulado...");

                    continue;
                }
                else if (diaSemana == 7) //rotina somente de sábado, simulando um dia mais "preguiçoso"
                {
                    time = getTimeSpan(time, new TimeSpan(10, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(11, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    //time = getTimeSpan(time, new TimeSpan(12.50, 0), rndMin, rndMax);
                    time = getTimeSpan(time, new TimeSpan(12, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(13, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(14, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    //time = getTimeSpan(time, new TimeSpan(18.00 + GetRandomDouble(-0.125, 0.125), 2));
                    time = getTimeSpan(time, new TimeSpan(18, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(19, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(20, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = getTimeSpan(time, new TimeSpan(01, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    Console.WriteLine("Dia " + i + " simulado...");

                    continue;
                }

                time = getTimeSpan(time, new TimeSpan(07, 15, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                time = getTimeSpan(time, new TimeSpan(08, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                time = getTimeSpan(time, new TimeSpan(08, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                time = getTimeSpan(time, new TimeSpan(09, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "F", diaSemana != 1 && diaSemana != 7));

                time = getTimeSpan(time, new TimeSpan(12, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                time = getTimeSpan(time, new TimeSpan(13, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                //time = getTimeSpan(time, new TimeSpan(13.50 + GetRandomDouble(-0.125, 0.125), 2));
                time = getTimeSpan(time, new TimeSpan(13, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "F", diaSemana != 1 && diaSemana != 7));

                time = getTimeSpan(time, new TimeSpan(17, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                time = getTimeSpan(time, new TimeSpan(19, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                //time = getTimeSpan(time, new TimeSpan(19.50 + GetRandomDouble(-0.125, 0.125), 2));
                time = getTimeSpan(time, new TimeSpan(19, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                time = getTimeSpan(time, new TimeSpan(20, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                time = getTimeSpan(time, new TimeSpan(22, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                time = getTimeSpan(time, new TimeSpan(23, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                Console.WriteLine("Dia " + i + " simulado...");

            }

            // Gerar anomalias
            //var countAnomalias = (int)(rotinas.Count * (percAnomalias / 100));
            var countAnomalias = (int)(dias * (percAnomalias / 100));
            int iAnomalias = 0;
            while (iAnomalias < countAnomalias)
            {
                var iDia = random.Next(0, dias);
                var dtAnomalia = dataBase.AddDays(iDia);
                var diaSemana = getDiaSemana(dtAnomalia.DayOfWeek);
                if (diaSemana != 1 && diaSemana != 7)
                {
                    if (random.NextDouble() > 0.9)
                    {
                        // ir ao banheiro de madrugada
                        var hr = getTimeSpan(TimeSpan.Zero, new TimeSpan(4, 30, 0), rndMin, rndMax);
                        rotinas.Add(new Rotina(dtAnomalia, hr, "B", diaSemana != 1 && diaSemana != 7));
                        iAnomalias++;
                    }
                    else if (random.NextDouble() > 0.9)
                    {
                        // ir ao banheiro ao meio-dia
                        var hr = getTimeSpan(TimeSpan.Zero, new TimeSpan(12, 30, 0), rndMin, rndMax);
                        rotinas.Add(new Rotina(dtAnomalia, hr, "B", diaSemana != 1 && diaSemana != 7));
                        iAnomalias++;
                    }
                    else if (random.NextDouble() > 0.9)
                    {
                        // chegar mais cedo em casa e ir ao banheiro
                        var hr = getTimeSpan(TimeSpan.Zero, new TimeSpan(15, 00, 0), rndMin, rndMax);
                        rotinas.Add(new Rotina(dtAnomalia, hr, "S", diaSemana != 1 && diaSemana != 7));
                        hr = getTimeSpan(hr, new TimeSpan(15, 30, 0), rndMin, rndMax);
                        rotinas.Add(new Rotina(dtAnomalia, hr, "B", diaSemana != 1 && diaSemana != 7));
                        hr = getTimeSpan(hr, new TimeSpan(15, 35, 0), rndMin, rndMax);
                        rotinas.Add(new Rotina(dtAnomalia, hr, "S", diaSemana != 1 && diaSemana != 7));
                        iAnomalias++;
                    }
                    else if (random.NextDouble() > 0.9)
                    {
                        // ir para fora de madrugada
                        var hr = getTimeSpan(TimeSpan.Zero, new TimeSpan(3, 0, 0), rndMin, rndMax);
                        rotinas.Add(new Rotina(dtAnomalia, hr, "F", diaSemana != 1 && diaSemana != 7));
                        iAnomalias++;
                    }

                }

            }

            foreach (var rotina in rotinas)
            {
                WriteLine(sb, rotina, algoritmo);
            }

            return iAnomalias;
        }


        private static int Rotina1(int dias, int nAnomalias, int rndMin, int rndMax, List<string> sb, int algoritmo)
        {
            var rotinas = new List<Rotina>();
            int cAnomalias = 0;
            var dataBase = new DateTime(2019, 1, 1);
            var time = new TimeSpan(0, 0, 0);

            for (var i = 0; i <= dias; i++)
            {
                //if (cAnomalias < nAnomalias && random.NextDouble() > 0.7)
                //{
                //    time = getTimeSpan(time, new TimeSpan(5, 30, 0), rndMin * 2, rndMax * 2);
                //    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "F", true));
                //    cAnomalias++;
                //}

                time = getTimeSpan(time, new TimeSpan(7, 15, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", true));

                time = getTimeSpan(time, new TimeSpan(8, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", true));

                time = getTimeSpan(time, new TimeSpan(8, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", true));

                time = getTimeSpan(time, new TimeSpan(9, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "F", true));

                time = getTimeSpan(time, new TimeSpan(12, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", true));

                time = getTimeSpan(time, new TimeSpan(13, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", true));

                time = getTimeSpan(time, new TimeSpan(13, 10, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "F", true));

                time = getTimeSpan(time, new TimeSpan(17, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", true));

                time = getTimeSpan(time, new TimeSpan(19, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", true));

                time = getTimeSpan(time, new TimeSpan(19, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", true));

                time = getTimeSpan(time, new TimeSpan(20, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", true));

                //if (cAnomalias < nAnomalias && random.NextDouble() > 0.7)
                //{
                //    time = getTimeSpan(time, new TimeSpan(21, 15, 0), rndMin, rndMax);
                //    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", true));
                //    cAnomalias++;

                //    time = getTimeSpan(time, new TimeSpan(21, 25, 0), rndMin, rndMax);
                //    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", true));
                //    cAnomalias++;
                //}


                time = getTimeSpan(time, new TimeSpan(22, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", true));

                time = getTimeSpan(time, new TimeSpan(23, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", true));


                Console.WriteLine("Dia " + i + " simulado...");
            }

            foreach (var rotina in rotinas)
            {
                WriteLine(sb, rotina, algoritmo);
            }

            return cAnomalias;
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

        private static TimeSpan getTimeSpan(TimeSpan lasTime, TimeSpan baseTime, int rndMin, int rndMax, TimeSpan? maxTime = null)
        {
            var diff = (int)(lasTime - baseTime).TotalMinutes;
            if (diff > 0)
            {
                diff = (int)(lasTime.Add(TimeSpan.FromDays(-1)) - baseTime).TotalMinutes;
            }

            var diff2 = int.MaxValue;
            if (maxTime.HasValue)
            {
                //diff2 = (int)(maxTime.Value - baseTime).TotalMinutes;
                //if (diff2 < 0)
                //{
                //    diff2 = (int)(maxTime.Value.Add(TimeSpan.FromDays(-1)) - baseTime).TotalMinutes;
                //}
                diff2 = (int)(maxTime.Value - baseTime).TotalMinutes;
            }

            var rndMinutes = random.Next(Math.Max(rndMin, diff), Math.Min(rndMax, diff2));
            var result = baseTime.Add(TimeSpan.FromMinutes(rndMinutes));
            return result;
        }

        private static void WriteLine(List<string> sb, Rotina rotina, int algoritmo)
        {
            var sLine = new List<string>();

            switch (algoritmo)
            {
                case 1: // DBSCAN
                case 2: // LOF
                    var tm = rotina.DiaMes.Add(rotina.Hora);
                    sLine.Add($"{tm.Hour}.{(tm.Minute / 60.0 * 100):00}");

                    sLine.Add((getDiaSemana(rotina.DiaMes.DayOfWeek) * 1000 + int.Parse(mapComodo[rotina.Comodo])).ToString());
                    break;

                //case 3: // OPTICS

                default:
                    sLine.Add(rotina.DiaMes.ToShortDateString());
                    sLine.Add(rotina.DiaMes.Add(rotina.Hora).ToShortTimeString());
                    sLine.Add(rotina.Comodo);
                    sLine.Add(rotina.isDiaUtil ? "Sim" : "Não");
                    break;
            }

            sb.Add(string.Join(separator, sLine));
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

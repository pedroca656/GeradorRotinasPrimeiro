using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace GeraRotina
{
    class Program
    {
        private static string separator = ",";
        private static Random random = new Random();

        // mapear cômodos para inteiro
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
                Console.WriteLine("Parâmetros inválidos!");
                Console.WriteLine("Usage: GeraRotina.exe numDias NomeArquivo numAlgoritmo [percAnomalias]\n" +
                                  "\tnumDia:      quantidade de dias a serem simulados.\n" +
                                  "\tnomeArquivo: nome do arquivo a ser gerado.\n" +
                                  "\numAlgoritmo: id do algoritmo a ser utilizado para teste (1 = DBSCAN; 2 = LOF; 3 = OPTICS).\n" +
                                  "\tpercAnomlais: porcentagem de anomalias a serem geradas.(opcional)\n" +
                                  "");
                return;
            }

            int dias;
            var percAnomalias = 0m;
            int algoritmo;

            // ***  Lê parametros   ***
            if (!int.TryParse(args[0], out dias))
            {
                Console.WriteLine("Parâmetro numDias incorreto.");
                throw new ArgumentException($"Parâmetro numDias incorreto.");
            }

            if (!int.TryParse(args[2], out algoritmo))
            {
                Console.WriteLine("Parâmetro idAlgoritmo incorreto.");
                throw new ArgumentException($"Parâmetro idAlgoritmo incorreto.");
            }

            if (args.Length >= 4 && !decimal.TryParse(args[3], out percAnomalias))
            {
                throw new ArgumentException($"Parâmetro percAnomalias incorreto.");
            }

            Console.WriteLine("Iniciando aplicação...");

            var sLines = new List<string>();

            // Cabeçalho
            var sCabecalho = new List<string>();
            switch (algoritmo)
            {
                case 1: // DBSCAN
                case 2: // LOF
                case 3: // OPTICS
                    sCabecalho.Add("Hora"); //hora será em números inteiros
                    sCabecalho.Add("DiaSemana-Comodo"); //comodo trocado para B = 100, Q = 200, C = 300, F = 400, S = 500 para poder executar o algoritmo
                    break;

                default:
                    sCabecalho.Add("Dia Mês");
                    sCabecalho.Add("Hora");
                    sCabecalho.Add("Cômodo");
                    sCabecalho.Add("Dia Útil?");

                    break;
            }
            sLines.Add(string.Join(separator, sCabecalho));

            // random default para execução de ações da rotina
            var rndMin = -15;
            var rndMax = 15;

            var anomaliasGeradas = Rotina2(dias, percAnomalias, rndMin, rndMax, sLines, algoritmo);

            // Criar diretório para salvar datasets
            var path = Path.Combine(AppContext.BaseDirectory, "output");
            Directory.CreateDirectory(path);

#if DEBUG
            // Gerar arquivos em pastas do google drive
            if (Environment.MachineName == "RODRIGO-NOTE")
            { path = Path.Combine(@"C:\Users\rodri\Google Drive\Datasets-beta"); }
            if (Environment.MachineName == "PEDRO-NOTE")
            { path = Path.Combine(@"C:\Users\pedro\Google Drive\Datasets-beta"); }
#endif

#if DEBUG
            var fileName = Path.Combine(path, args[1] + $"__{anomaliasGeradas}__" + ".csv");
#else
            var fileName = Path.Combine(path, args[1] + ".csv");
#endif
            //  *** Gerar arquivo .csv  ***
            using (var fs = new StreamWriter(fileName))
            {
                fs.Write(string.Join("\n", sLines));
            }

            Console.WriteLine($"Arquivo gerado! {fileName}");

#if DEBUG
            //  *** Abrir arquivo   ***
            System.Diagnostics.Process.Start(fileName);
#endif
        }

        private static int Rotina2(int dias, decimal percAnomalias, int rndMin, int rndMax, List<string> sb, int algoritmo)
        {
            var rotinas = new List<Rotina>();
            var dataBase = new DateTime(2019, 1, 1);
            var time = new TimeSpan(0, 0, 0);

            // ***  Gerar rotina    ***

            for (int i = 0; i <= dias; i++)
            {
                var diaSemana = GetDiaSemana(dataBase.DayOfWeek);
                if (diaSemana == 4) //rotina somente de quarta, simulando um dia em que o indivíduo chega mais tarde em casa
                {
                    time = GetTimeSpan(time, new TimeSpan(7, 10, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", true));

                    time = GetTimeSpan(time, new TimeSpan(08, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(08, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(09, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "F", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(12, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(13, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    //time = getTimeSpan(time, new TimeSpan(13.50 + GetRandomDouble(-0.125, 0.125), 2));
                    time = GetTimeSpan(time, new TimeSpan(13, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "F", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(22, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(23, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(23, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    Console.WriteLine("Dia " + i + " simulado...");

                    continue;
                }
                else if (diaSemana == 1) //rotina somente de domingo, simulando um almoço em família
                {
                    time = GetTimeSpan(time, new TimeSpan(09, 15, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(09, 50, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(10, 15, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(12, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(14, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    //time = getTimeSpan(time, new TimeSpan(16.00 + GetRandomDouble(-0.125, 0.125), 2));
                    time = GetTimeSpan(time, new TimeSpan(16, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(19, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(19, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(20, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(23, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    Console.WriteLine("Dia " + i + " simulado...");

                    continue;
                }
                else if (diaSemana == 7) //rotina somente de sábado, simulando um dia mais "preguiçoso"
                {
                    time = GetTimeSpan(time, new TimeSpan(10, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(11, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    //time = getTimeSpan(time, new TimeSpan(12.50, 0), rndMin, rndMax);
                    time = GetTimeSpan(time, new TimeSpan(12, 30, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(13, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(14, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    //time = getTimeSpan(time, new TimeSpan(18.00 + GetRandomDouble(-0.125, 0.125), 2));
                    time = GetTimeSpan(time, new TimeSpan(18, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(19, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(20, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                    time = GetTimeSpan(time, new TimeSpan(01, 0, 0), rndMin, rndMax);
                    rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                    Console.WriteLine("Dia " + i + " simulado...");

                    continue;
                }

                time = GetTimeSpan(time, new TimeSpan(07, 15, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                time = GetTimeSpan(time, new TimeSpan(08, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                time = GetTimeSpan(time, new TimeSpan(08, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                time = GetTimeSpan(time, new TimeSpan(09, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "F", diaSemana != 1 && diaSemana != 7));

                time = GetTimeSpan(time, new TimeSpan(12, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                time = GetTimeSpan(time, new TimeSpan(13, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                //time = getTimeSpan(time, new TimeSpan(13.50 + GetRandomDouble(-0.125, 0.125), 2));
                time = GetTimeSpan(time, new TimeSpan(13, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "F", diaSemana != 1 && diaSemana != 7));

                time = GetTimeSpan(time, new TimeSpan(17, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                time = GetTimeSpan(time, new TimeSpan(19, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "B", diaSemana != 1 && diaSemana != 7));

                //time = getTimeSpan(time, new TimeSpan(19.50 + GetRandomDouble(-0.125, 0.125), 2));
                time = GetTimeSpan(time, new TimeSpan(19, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                time = GetTimeSpan(time, new TimeSpan(20, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "C", diaSemana != 1 && diaSemana != 7));

                time = GetTimeSpan(time, new TimeSpan(22, 0, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "S", diaSemana != 1 && diaSemana != 7));

                time = GetTimeSpan(time, new TimeSpan(23, 30, 0), rndMin, rndMax);
                rotinas.Add(new Rotina(dataBase.AddDays(i), time, "Q", diaSemana != 1 && diaSemana != 7));

                Console.WriteLine("Dia " + i + " simulado...");

            }

            
            // ***  Gerar anomalias ***

            var countAnomalias = (int)(rotinas.Count * (percAnomalias / 100));
            int iAnomalias = 0;
            while (iAnomalias < countAnomalias)
            {
                var iDia = random.Next(0, dias);
                var dtAnomalia = dataBase.AddDays(iDia);
                var diaSemana = GetDiaSemana(dtAnomalia.DayOfWeek);
                if (diaSemana != 1 && diaSemana != 7)
                {
                    if (random.NextDouble() > 0.9)
                    {
                        // ir ao banheiro de madrugada
                        var hr = GetTimeSpan(TimeSpan.Zero, new TimeSpan(4, 30, 0), rndMin, rndMax);
                        rotinas.Add(new Rotina(dtAnomalia, hr, "B", diaSemana != 1 && diaSemana != 7));
                        iAnomalias++;
                    }
                    else if (random.NextDouble() > 0.9)
                    {
                        // ir ao banheiro ao meio-dia
                        var hr = GetTimeSpan(TimeSpan.Zero, new TimeSpan(12, 30, 0), rndMin, rndMax);
                        rotinas.Add(new Rotina(dtAnomalia, hr, "B", diaSemana != 1 && diaSemana != 7));
                        iAnomalias++;
                    }
                    else if (random.NextDouble() > 0.9)
                    {
                        // chegar mais cedo em casa e ir ao banheiro
                        var hr = GetTimeSpan(TimeSpan.Zero, new TimeSpan(15, 00, 0), rndMin, rndMax);
                        rotinas.Add(new Rotina(dtAnomalia, hr, "S", diaSemana != 1 && diaSemana != 7));
                        iAnomalias++;
                        hr = GetTimeSpan(hr, new TimeSpan(15, 30, 0), rndMin, rndMax);
                        rotinas.Add(new Rotina(dtAnomalia, hr, "B", diaSemana != 1 && diaSemana != 7));
                        iAnomalias++;
                        hr = GetTimeSpan(hr, new TimeSpan(15, 35, 0), rndMin, rndMax);
                        rotinas.Add(new Rotina(dtAnomalia, hr, "S", diaSemana != 1 && diaSemana != 7));
                        iAnomalias++;
                    }
                    else if (random.NextDouble() > 0.9)
                    {
                        // ir para fora de casa na madrugada
                        var hr = GetTimeSpan(TimeSpan.Zero, new TimeSpan(3, 0, 0), rndMin, rndMax);
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


        private static int GetDiaSemana(DayOfWeek dia)
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

        private static TimeSpan GetTimeSpan(TimeSpan lasTime, TimeSpan baseTime, int rndMin, int rndMax, TimeSpan? maxTime = null)
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
                case 3: // OPTICS
                    var tm = rotina.DiaMes.Add(rotina.Hora);
                    sLine.Add($"{tm.Hour}.{(tm.Minute / 60.0 * 100):00}");

                    sLine.Add((GetDiaSemana(rotina.DiaMes.DayOfWeek) * 1000 + int.Parse(mapComodo[rotina.Comodo])).ToString());
                    break;

                default:
                    sLine.Add(rotina.DiaMes.ToShortDateString());
                    sLine.Add(rotina.DiaMes.Add(rotina.Hora).ToShortTimeString());
                    sLine.Add(rotina.Comodo);
                    sLine.Add(rotina.IsDiaUtil ? "Sim" : "Não");
                    break;
            }

            sb.Add(string.Join(separator, sLine));
        }

        private class Rotina
        {
            public DateTime DiaMes { get; set; }
            public TimeSpan Hora { get; set; }
            public string Comodo { get; set; }
            public bool IsDiaUtil { get; set; }

            public Rotina(DateTime diaMes, TimeSpan hora, string comodo, bool isDiaUtil)
            {
                DiaMes = diaMes;
                Hora = hora;
                Comodo = comodo;
                this.IsDiaUtil = isDiaUtil;
            }
        }
    }
}

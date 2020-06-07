using System;
using System.Collections.Generic;

namespace Mao.Relatorios.Classes
{
    public class Frequencia
    {
        public DateTime Data { get; set; }
        public TimeSpan Entrada { get; set; }
        public TimeSpan Saida { get; set; }
        public TimeSpan TotalHorasDia { get; set; }
        public string Tarefa { get; set; }


        public static List<Frequencia> GetListaFrequencia()
        {
            return new List<Frequencia>()
            {
                new Frequencia() { Data = Convert.ToDateTime("14/04/2020"), Entrada = TimeSpan.Parse("07:00:00"), Saida = TimeSpan.Parse("13:00:00"), TotalHorasDia = TimeSpan.Parse("06:00:00"), Tarefa = "Varrer" },
                new Frequencia() { Data = Convert.ToDateTime("14/04/2020"), Entrada = TimeSpan.Parse("20:00:00"), Saida = TimeSpan.Parse("21:01:00"), TotalHorasDia = TimeSpan.Parse("01:01:00"), Tarefa = "Segurança" },
                new Frequencia() { Data = Convert.ToDateTime("15/04/2020"), Entrada = TimeSpan.Parse("12:30:00"), Saida = TimeSpan.Parse("13:00:00"), TotalHorasDia = TimeSpan.Parse("00:30:00"), Tarefa = "Lava louça" },
                new Frequencia() { Data = Convert.ToDateTime("25/04/2020"), Entrada = TimeSpan.Parse("19:00:00"), Saida = TimeSpan.Parse("21:30:00"), TotalHorasDia = TimeSpan.Parse("02:30:00"), Tarefa = "Cozinhar" },
                new Frequencia() { Data = Convert.ToDateTime("06/05/2020"), Entrada = TimeSpan.Parse("10:00:00"), Saida = TimeSpan.Parse("13:30:00"), TotalHorasDia = TimeSpan.Parse("03:30:00"), Tarefa = "Monitorar alunos" },
                new Frequencia() { Data = Convert.ToDateTime("07/05/2020"), Entrada = TimeSpan.Parse("09:00:00"), Saida = TimeSpan.Parse("10:00:00"), TotalHorasDia = TimeSpan.Parse("01:00:00"), Tarefa = "Pintura" },
                new Frequencia() { Data = Convert.ToDateTime("14/05/2020"), Entrada = TimeSpan.Parse("10:45:00"), Saida = TimeSpan.Parse("12:00:00"), TotalHorasDia = TimeSpan.Parse("01:15:00"), Tarefa = "Fazer merenda" }
            };
        }
    }
}

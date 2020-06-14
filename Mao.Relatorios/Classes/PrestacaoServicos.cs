using System;

namespace Mao.Relatorios.Classes
{
    public class PrestacaoServicos
    {
        public string NomeEscola { get; set; }
        public string EnderecoEscola { get; set; }
        public string TelefoneEscola { get; set; }
        public string CepEscola { get; set; }

        public string Nome { get; set; }
        public string Rg { get; set; }
        public string Cpf { get; set; }
        public string Naturalidade { get; set; }
        public string Idade { get; set; }
        public string EstadoCivil { get; set; }
        public string Escolaridade { get; set; }
        public string Residencia { get; set; }
        public string Telefone { get; set; }
        public string Profissao { get; set; }
        public string OrgaoOrigem { get; set; }
        public string MotivoPrestacaoServico { get; set; }
        public string PrazoPrestacaoServico { get; set; }
        public DateTime DataApresentacao { get; set; }


        public static PrestacaoServicos GetPrestacaoServicos()
        {
            return new PrestacaoServicos()
            {
                NomeEscola = "E.E. Prof. Isaltino de Mello",
                EnderecoEscola = "Rua Toninhas, 73 - Campo Grande",
                TelefoneEscola = "(11) 5631-7044",
                CepEscola = "04679-540",

                Nome = "Bushmaster Francisco",
                Rg = "41.875.789-6",
                Cpf = "030.540.898-43",
                Naturalidade = "Jamaicana",
                Idade = "35",
                EstadoCivil = "Casado",
                Escolaridade = "Superior Completo",
                Residencia = "Rua Dr. Mario Rocco, 104",
                Telefone = "(11) 95631-3256",
                Profissao = "analista de sistemas",
                OrgaoOrigem = "Prefeitura Municipal de Santo Amaro",
                MotivoPrestacaoServico = "Vestibulum ante ipsum primis in faucibus orci luctus et ultrices posuere cubilia Curae.",
                PrazoPrestacaoServico = "45 dias corridos",
                DataApresentacao = new DateTime(2020, 6, 20, 08, 10, 20)
            };
        }
    }
}
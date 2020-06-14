using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Mao.Relatorios.Classes;
using Mao.Relatorios.Core.PDF;

namespace Mao.Relatorios.Reports
{
    public class RelatorioPrestacaoServico
    {
        private static PdfConverter GetConfigurations()
        {
            PdfConverter pdfConverter = new PdfConverter();
            pdfConverter.DocumentOptions.PageOrientation = PdfPageOrientation.Retrato;
            pdfConverter.DocumentOptions.ShowFooter = true;
            pdfConverter.DocumentOptions.ShowHeader = true;
            pdfConverter.DocumentOptions.ShowPagingOnBottom = true;
            pdfConverter.DocumentOptions.ShowPrintDateTime = false;
            pdfConverter.DocumentOptions.HeaderOptions.DrawHeaderLine = true;
            pdfConverter.DocumentOptions.HeaderOptions.HeaderTitleText = "";
            pdfConverter.DocumentOptions.HeaderOptions.HeaderImageLeft = System.Drawing.Image.FromFile($"{Directory.GetCurrentDirectory()}\\images\\fde_logo_2.jpg");
            pdfConverter.DocumentOptions.HeaderOptions.HeaderImageRight = System.Drawing.Image.FromFile($"{Directory.GetCurrentDirectory()}\\images\\gov_sp_logo.png");
            pdfConverter.DocumentOptions.HeaderOptions.ShowImageRight = true;
            pdfConverter.DocumentOptions.FooterOptions.DrawFooterLine = false;
            pdfConverter.DocumentOptions.FooterOptions.ShowFooterText = true;
            pdfConverter.DocumentOptions.FooterOptions.ShowPageNumber = true;
            pdfConverter.DocumentOptions.DocumentInfo.AuthorName = "Rafael Francisco";
            pdfConverter.DocumentOptions.DocumentInfo.CreatedDate = DateTime.Now;
            pdfConverter.DocumentOptions.DocumentInfo.Title = "Relatório de Prestação de Contas";
            pdfConverter.DocumentOptions.DocumentInfo.Subject = "Relatório gerado para testes";
            pdfConverter.DocumentOptions.DocumentInfo.Creator = "Sistema de teste de geração de PDF";

            return pdfConverter;
        }
        private static string GetHtmlContent(PrestacaoServicos dados)
        {
            StringBuilder html = new StringBuilder(PdfMethods.LoadHtmlTemplate($"{Directory.GetCurrentDirectory()}\\templates_relatorios\\RelatorioPrestacaoServico_Template.html"));

            html = html.Replace("#KEY_NOME_ESCOLA", dados.NomeEscola)
                       .Replace("#KEY_ENDERECO_ESCOLA", dados.EnderecoEscola);

            return html.ToString();
        }

        public static byte[] Generate(PrestacaoServicos dados)
        {
            PdfConverter pdf = GetConfigurations();
            return pdf.RelatorioPrestacaoServico(dados);
        }
        public static byte[] GenerateFromHtml(PrestacaoServicos dados)
        {
            string html = GetHtmlContent(dados);
            PdfConverter pdf = GetConfigurations();
            return pdf.GeneratePdfBytesFromHtmlString(html);
        }
    }
}

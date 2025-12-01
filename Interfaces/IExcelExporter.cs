public interface IExcelExporter
{
    void ExportarCotacoesParaExcel(IEnumerable<CotacaoDto> cotacoes, string caminhoArquivo);
}

public class Cotacao
{
    public int CotacaoId { get; set; }
    public int LoteCotacaoId { get; set; }
    public DateTime DataCotacao { get; set; }
    public string TipoMoeda { get; set; } = string.Empty;
    public decimal ValorCompra { get; set; }
    public decimal ValorVenda { get; set; }
    public DateTime DataInclusao { get; set; }
}
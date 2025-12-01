public class ParametroColeta
{
    public int ParametroColetaId { get; set; }
    public DateTime DataInicial { get; set; }
    public DateTime DataFinal { get; set; }
    public string TipoMoeda { get; set; } = string.Empty;
    public bool Ativo { get; set; }
    public DateTime DataInclusao { get; set; }
}
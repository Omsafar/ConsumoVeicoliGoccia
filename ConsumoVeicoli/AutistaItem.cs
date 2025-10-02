namespace ConsumoVeicoli.Models
{
    /// <summary>Elemento da mostrare nella ListBox quando si lavora per autisti.</summary>
    internal sealed class AutistaItem
    {
        public string Codice { get; set; } = "";
        public string Nome { get; set; } = "";

        public override string ToString() =>
            string.IsNullOrWhiteSpace(Nome) ? Codice : $"{Codice} – {Nome}";
    }
}

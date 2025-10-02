namespace ConsumoVeicoli.Models
{
    public class DatiConsumoRecord
    {
        public string Targa { get; set; } = "";
        public string Numero_Interno { get; set; } = "";
        public System.DateTime Data { get; set; }
        public int Km_Totali { get; set; }
        public decimal Litri_Totali { get; set; }
        public decimal ConsumoKmPerLitro { get; set; }
    }
}

namespace ConsumoVeicoli.Models
{
    public class ConsumoMedioResult
    {
        public string Targa { get; set; } = "";
        // Media esistente (giornaliera), calcolata da tbDatiConsumo
        public decimal MediaConsumo { get; set; }
        // Nuova colonna "Km/L Rifornimento"
        public decimal MediaConsumoRifornimenti { get; set; }
        public string MediaConsumoRifornimentiText
        {
            get
            {
                if (MediaConsumoRifornimenti < 0)
                    return "veicolo a metano";
                else
                    return MediaConsumoRifornimenti.ToString("F2");
            }
        }
    }
}
using System;

namespace ConsumoVeicoli.Models
{
    public class DatiConsumoAutistaRecord
    {
        public string CodiceAutista { get; set; } = "";
        public string NomeAutista { get; set; } = "";
        public DateTime DataGiorno { get; set; }
        public long KmPercorsi { get; set; }
        public decimal LitriConsumati { get; set; }
        public decimal KmPerLitro { get; set; }
        public string Veicoli { get; set; } = "";
    }
}

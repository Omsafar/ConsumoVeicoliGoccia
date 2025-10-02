namespace ConsumoVeicoli.Models
{
    public class RifornimentoRecord
    {
        public string Veicolo { get; set; } = "";
        public System.DateTime DataOra { get; set; }
        public decimal Litri { get; set; }

        public decimal Kg { get; set; }
    }
}

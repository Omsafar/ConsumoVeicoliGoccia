using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using Microsoft.Data.SqlClient;
using Microsoft.Win32;
using ClosedXML.Excel;
using ConsumoVeicoli.Models;

namespace ConsumoVeicoli
{
    public partial class MainWindow : Window
    {
        private List<string> _allTarghe = new();
        private List<AutistaItem> _allAutisti = new();
        private readonly HashSet<string> _selectedTarghe = new();
        private readonly HashSet<string> _selectedAutisti = new();
        private const string ConnString =
            "Server=srv2016app02\\sgam;Database=PARATORI;User Id=sapara;Password=S@p4ra;Encrypt=True;TrustServerCertificate=True;";

        private const string ConnStringSgam =
            "Server=srv2016app02\\sgam;Database=SGAM;User Id=sapara;Password=S@p4ra;Encrypt=True;TrustServerCertificate=True;";

        private enum ModoRicerca { Targhe, Autisti }
        private ModoRicerca _modo = ModoRicerca.Targhe;

        private Dictionary<string, string> _mappaCarburante = new Dictionary<string, string>(); // per i veicoli

        private bool _isRefreshingList;

        public MainWindow()
        {
            InitializeComponent();
            lbRisorse.SelectionChanged += LbRisorse_SelectionChanged;
            cbModoRicerca.SelectedIndex = 0;    // qui è sicuro che cbModoRicerca sia != null
            CaricaTarghe();
            CostruisciColonneVeicoli();
        }

        #region Caricamento risorse (targhe / autisti)

        private void CaricaTarghe()
        {
            _allTarghe.Clear();
            using var cn = new SqlConnection(ConnString);
            cn.Open();
            var cmd = new SqlCommand(
                "SELECT DISTINCT Targa FROM dbo.tbDatiConsumo WHERE Targa IS NOT NULL ORDER BY Targa", cn);
            using var dr = cmd.ExecuteReader();
            while (dr.Read())
                _allTarghe.Add(dr.GetString(0));

            PopolaTarghe(_allTarghe);
        }

        private void CaricaAutisti()
        {
            _allAutisti.Clear();
            using var cn = new SqlConnection(ConnString);
            cn.Open();
            var cmd = new SqlCommand(@"
        SELECT DISTINCT CodiceAutista, NomeAutista
        FROM dbo.tbDatiConsumoAutisti
        WHERE CodiceAutista IS NOT NULL
        ORDER BY CodiceAutista", cn);
            using var dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                _allAutisti.Add(new AutistaItem
                {
                    Codice = dr["CodiceAutista"] as string ?? "",
                    Nome = dr["NomeAutista"] as string ?? ""
                });
            }

            PopolaAutisti(_allAutisti);
        }


        #endregion

        #region Gestione UI

        private void cbModoRicerca_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            _modo = cbModoRicerca.SelectedIndex == 0 ? ModoRicerca.Targhe : ModoRicerca.Autisti;

            if (_modo == ModoRicerca.Targhe)
            {
                lblRisorsa.Text = "Targhe:";
                CaricaTarghe();
                CostruisciColonneVeicoli();
            }
            else
            {
                lblRisorsa.Text = "Autisti:";
                CaricaAutisti();
                CostruisciColonneAutisti();
            }

            // ripristino stato pulsante seleziona tutte
            btnToggleAll.Content = "Seleziona Tutte";
        }

        private void btnToggleAll_Click(object sender, RoutedEventArgs e)
        {
            bool seleziona = lbRisorse.SelectedItems.Count < lbRisorse.Items.Count;

            if (_modo == ModoRicerca.Targhe)
            {
                var visibles = lbRisorse.Items.Cast<string>();
                if (seleziona)
                    _selectedTarghe.UnionWith(visibles);
                else
                    foreach (var v in visibles) _selectedTarghe.Remove(v);
            }
            else
            {
                var visibles = lbRisorse.Items.Cast<AutistaItem>()
                                              .Select(a => a.Codice);
                if (seleziona)
                    _selectedAutisti.UnionWith(visibles);
                else
                    foreach (var v in visibles) _selectedAutisti.Remove(v);
            }

            if (seleziona)
            {
                lbRisorse.SelectAll();
                btnToggleAll.Content = "Deseleziona Tutte";
            }
            else
            {
                lbRisorse.UnselectAll();
                btnToggleAll.Content = "Seleziona Tutte";
            }
        }


        private void CostruisciColonneVeicoli()
        {
            dgRisultati.Columns.Clear();

            dgRisultati.Columns.Add(new DataGridTextColumn
            {
                Header = "Targa",
                Binding = new Binding("Targa"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                IsReadOnly = true
            });
            dgRisultati.Columns.Add(new DataGridTextColumn
            {
                Header = "Media Consumo (km/l)",
                Binding = new Binding("MediaConsumo"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                IsReadOnly = true
            });
            dgRisultati.Columns.Add(new DataGridTextColumn
            {
                Header = "Media Rifornimenti (km/l)",
                Binding = new Binding("MediaConsumoRifornimenti"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                IsReadOnly = true
            });
        }

        private void CostruisciColonneAutisti()
        {
            dgRisultati.Columns.Clear();

            dgRisultati.Columns.Add(new DataGridTextColumn
            {
                Header = "Codice Autista",
                Binding = new Binding("CodiceAutista"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                IsReadOnly = true
            });
            dgRisultati.Columns.Add(new DataGridTextColumn
            {
                Header = "Nome",
                Binding = new Binding("NomeAutista"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                IsReadOnly = true
            });
            dgRisultati.Columns.Add(new DataGridTextColumn
            {
                Header = "Media Consumo (km/l)",
                Binding = new Binding("MediaConsumo"),
                Width = new DataGridLength(1, DataGridLengthUnitType.Star),
                IsReadOnly = true
            });
        }

        #endregion

        #region Pulsante Carica

        private void btnCarica_Click(object sender, RoutedEventArgs e)
        {
            if (lbRisorse.SelectedItems == null || lbRisorse.SelectedItems.Count == 0)
            {
                MessageBox.Show("Seleziona almeno un elemento.");
                return;
            }

            if (_modo == ModoRicerca.Targhe)
                CaricaDatiVeicoli();
            else
                CaricaDatiAutisti();
        }

        private void CaricaDatiVeicoli()
        {
            if (rbStessoPeriodo.IsChecked == true)
            {
                var finestra = new FinestraDate();
                if (finestra.ShowDialog() != true) return;

                var da = finestra.DataDa;
                var a = finestra.DataA;
                if (da > a) { MessageBox.Show("Data Da > Data A."); return; }

                CaricaDatiStessoPeriodoVeicoli(lbRisorse.SelectedItems, da, a);
            }
            else
            {
                var lista = new List<PeriodoPerVeicolo>();
                foreach (var tObj in lbRisorse.SelectedItems)
                {
                    lista.Add(new PeriodoPerVeicolo
                    {
                        Targa = tObj?.ToString() ?? "",
                        DataDa = DateTime.Today,
                        DataA = DateTime.Today
                    });
                }

                var wnd = new PeriodiWindow(lista);
                if (wnd.ShowDialog() == true)
                    CaricaDatiPeriodiDiversiVeicoli(wnd.Risultato);
            }
        }

        private void CaricaDatiAutisti()
        {
            if (rbStessoPeriodo.IsChecked == true)
            {
                var finestra = new FinestraDate();
                if (finestra.ShowDialog() != true) return;

                var da = finestra.DataDa;
                var a = finestra.DataA;
                if (da > a) { MessageBox.Show("Data Da > Data A."); return; }

                CaricaDatiStessoPeriodoAutisti(lbRisorse.SelectedItems, da, a);
            }
            else
            {
                var lista = new List<PeriodoPerVeicolo>();
                foreach (AutistaItem aItem in lbRisorse.SelectedItems)
                {
                    lista.Add(new PeriodoPerVeicolo
                    {
                        Targa = aItem.Codice,      // riutilizziamo la classe esistente
                        DataDa = DateTime.Today,
                        DataA = DateTime.Today
                    });
                }

                var wnd = new PeriodiWindow(lista);
                if (wnd.ShowDialog() == true)
                    CaricaDatiPeriodiDiversiAutisti(wnd.Risultato);
            }
        }

        #endregion

        #region --- VEICOLI (codice già esistente, inalterato) ---

        private void CaricaDatiStessoPeriodoVeicoli(System.Collections.IList targhe, DateTime da, DateTime a)
        {
            if (_mappaCarburante == null || _mappaCarburante.Count == 0)
            {
                _mappaCarburante = CaricaTipiCarburante();
            }
            var risultati = new List<ConsumoMedioResult>();
            var messaggiCaricamento = new List<string>();

            foreach (var targaObj in targhe)
            {
                string targa = targaObj?.ToString() ?? "";
                string tipoCarb;
                if (_mappaCarburante.TryGetValue(targa, out tipoCarb!))
                    tipoCarb = tipoCarb?.ToUpper().Trim() ?? "GA";
                else
                    tipoCarb = "GA";
                var recs = LeggiDatiDaDb_Paratori(targa, da, a, messaggiCaricamento);
                var mediaGiornaliera = CalcolaMediaAritmeticaGiornaliera(recs);
                int sommaKm = recs.Sum(x => x.Km_Totali);
                Debug.WriteLine($"La somma di tutti i km percorsi per la targa {targa} è: {sommaKm}");
                var rifornimenti = LeggiRifornimentiDaDb_Sgam(
                    targa,
                    da,
                    a,
                    recs.Select(r => r.Numero_Interno));
                bool isMetano = (tipoCarb == "ME");
                // Calcolo la media usando KG se è metano, Litri altrimenti
                var mediaRifornimenti = CalcolaMediaSommaRifornimenti(recs, rifornimenti, isMetano);

                risultati.Add(new ConsumoMedioResult
                {
                    Targa = targa,
                    MediaConsumo = mediaGiornaliera,
                    MediaConsumoRifornimenti = mediaRifornimenti
                });
            }
            dgRisultati.ItemsSource = risultati;

            if (messaggiCaricamento.Count > 0)
            {
                MessageBox.Show(string.Join(Environment.NewLine, messaggiCaricamento));
            }
        }


        private void CaricaDatiPeriodiDiversiVeicoli(List<PeriodoPerVeicolo> lista)
        {
            // Se la mappa carburante non è stata caricata, la carico (usa la tua stored procedure)
            if (_mappaCarburante == null || _mappaCarburante.Count == 0)
            {
                _mappaCarburante = CaricaTipiCarburante();
            }

            // Dizionari per raggruppare i dati di consumo e i rifornimenti per ciascuna targa
            var datiMap = new Dictionary<string, List<DatiConsumoRecord>>();
            var rifornimentiMap = new Dictionary<string, List<RifornimentoRecord>>();

            // 1) Raccolgo i record per ciascun periodo/targa
            var messaggiCaricamento = new List<string>();


            foreach (var period in lista)
            {
                // Se non c'è ancora la chiave per questa targa, la inizializzo
                if (!datiMap.ContainsKey(period.Targa))
                {
                    datiMap[period.Targa] = new List<DatiConsumoRecord>();
                    rifornimentiMap[period.Targa] = new List<RifornimentoRecord>();
                }

                // Leggo i record di tbDatiConsumo per la (targa, DataDa, DataA)
                var recs = LeggiDatiDaDb_Paratori(period.Targa, period.DataDa, period.DataA, messaggiCaricamento);
                datiMap[period.Targa].AddRange(recs);

                // Leggo i rifornimenti per la (targa, DataDa, DataA)
                var rifs = LeggiRifornimentiDaDb_Sgam(
                    period.Targa,
                    period.DataDa,
                    period.DataA,
                    recs.Select(r => r.Numero_Interno));
                rifornimentiMap[period.Targa].AddRange(rifs);
            }

            // 2) Calcolo i risultati finali per ciascuna targa
            var risultati = new List<ConsumoMedioResult>();

            foreach (var targa in datiMap.Keys)
            {
                // a) Media dei consumi giornalieri (km/l) da tbDatiConsumo
                var recs = datiMap[targa];
                var mediaGiornaliera = CalcolaMediaAritmeticaGiornaliera(recs);

                // b) Verifico se il veicolo è a metano (ME) o gasolio (GA)
                string tipoCarb;
                if (_mappaCarburante.TryGetValue(targa, out tipoCarb!))
                    tipoCarb = tipoCarb?.ToUpper().Trim() ?? "";
                else
                    tipoCarb = "GA";  // default


                // c) Calcolo la media dei rifornimenti
                var rifornimenti = rifornimentiMap[targa];
                bool isMetano = (tipoCarb == "ME");

                // Se è metano => km/kg; se è gasolio => km/l
                var mediaRifornimenti = CalcolaMediaSommaRifornimenti(recs, rifornimenti, isMetano);

                // d) Infine aggiungo il risultato
                risultati.Add(new ConsumoMedioResult
                {
                    Targa = targa,
                    MediaConsumo = mediaGiornaliera,                // media giornaliera
                    MediaConsumoRifornimenti = mediaRifornimenti    // km/l oppure km/kg
                });
            }

            dgRisultati.ItemsSource = risultati;

            if (messaggiCaricamento.Count > 0)
            {
                MessageBox.Show(string.Join(Environment.NewLine, messaggiCaricamento));
            }
        }


        private List<DatiConsumoRecord> LeggiDatiDaDb_Paratori(
                string targa,
                DateTime d1,
                DateTime d2,
                List<string>? messaggiCaricamento = null)
        {
            var lista = new List<DatiConsumoRecord>();
            using var cn = new SqlConnection(ConnString);
            cn.Open();
            var cmd = new SqlCommand(@"
        SELECT Targa, Numero_Interno, Data, Km_Totali, Litri_Totali, [Consumo_km/l]
        FROM dbo.tbDatiConsumo
        WHERE Targa = @t AND Data >= @d1 AND Data <= @d2
        ORDER BY Data", cn);

            cmd.Parameters.AddWithValue("@t", targa);
            cmd.Parameters.AddWithValue("@d1", d1.Date);
            cmd.Parameters.AddWithValue("@d2", d2.Date.AddDays(1).AddTicks(-1));


            using var dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                var rec = new DatiConsumoRecord
                {
                    Targa = dr["Targa"] is DBNull ? "" : (string)dr["Targa"],
                    Numero_Interno = dr["Numero_Interno"] is DBNull ? "" : (string)dr["Numero_Interno"],
                    Data = dr["Data"] is DBNull ? DateTime.MinValue : (DateTime)dr["Data"],
                    Km_Totali = dr["Km_Totali"] is DBNull ? 0 : (int)dr["Km_Totali"],
                    Litri_Totali = dr["Litri_Totali"] is DBNull ? 0 : (decimal)dr["Litri_Totali"],
                    // mappiamo la nuova colonna Consumo_km/l:
                    ConsumoKmPerLitro = dr["Consumo_km/l"] is DBNull ? 0 : (decimal)dr["Consumo_km/l"]
                };
                lista.Add(rec);
            }
            if (messaggiCaricamento != null)
            {
                messaggiCaricamento.Add(
                    $"Per la targa {targa} dal {d1:dd/MM/yyyy} al {d2:dd/MM/yyyy} ho trovato {lista.Count} record");
            }
            return lista;
        }



        private List<RifornimentoRecord> LeggiRifornimentiDaDb_Sgam(
            string targa,
            DateTime d1,
            DateTime d2,
            IEnumerable<string>? codiciAlternativi = null)
        {
            var lista = new List<RifornimentoRecord>();
            using var cn = new SqlConnection(ConnStringSgam);
            var codiciDaCercare = new List<string>();
            var visti = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            void AggiungiCodice(string? codice)
            {
                if (string.IsNullOrWhiteSpace(codice))
                    return;

                // Il campo VEICOLO in SGAM accetta solo il numero interno senza spazi.
                // Rimuoviamo quindi tutti i caratteri di spazio e tabulazione prima di eseguire la query.
                var codiceNormalizzato = new string(codice.Where(c => !char.IsWhiteSpace(c)).ToArray());
                if (string.IsNullOrEmpty(codiceNormalizzato))
                    return;

                if (visti.Add(codiceNormalizzato))
                    codiciDaCercare.Add(codiceNormalizzato);
            }

            AggiungiCodice(targa);

            if (codiciAlternativi != null)
            {
                foreach (var codice in codiciAlternativi)
                    AggiungiCodice(codice);
            }

            if (codiciDaCercare.Count == 0)
                return lista;

            cn.Open();

            foreach (var codice in codiciDaCercare)
            {
                Debug.WriteLine($"Cerco rifornimenti per VEICOLO='{codice}' con data >= {d1} e data <= {d2}");
                using var cmd = new SqlCommand(@"
        SELECT VEICOLO, DATA_ORA, LITRI, KG
        FROM dbo.RisorseRifornimentiRilevazioni
        WHERE LTRIM(RTRIM(VEICOLO)) = @t
          AND DATA_ORA >= @d1
          AND DATA_ORA <= @d2
        ORDER BY DATA_ORA", cn);

                cmd.Parameters.AddWithValue("@t", codice);
                cmd.Parameters.AddWithValue("@d1", d1.Date);
                cmd.Parameters.AddWithValue("@d2", d2.Date.AddDays(1).AddTicks(-1));

                using var dr = cmd.ExecuteReader();
                while (dr.Read())
                {
                    var r = new RifornimentoRecord
                    {
                        Veicolo = dr["VEICOLO"] is DBNull ? "" : (string)dr["VEICOLO"],
                        DataOra = dr["DATA_ORA"] is DBNull ? DateTime.MinValue : (DateTime)dr["DATA_ORA"],
                        Litri = dr["LITRI"] is DBNull ? 0 : (decimal)dr["LITRI"],
                        Kg = dr["KG"] is DBNull ? 0 : (decimal)dr["KG"]  // Leggo anche i KG
                    };
                    Debug.WriteLine($"[SGAM] Trovato rifornimento: {r.Veicolo}, {r.DataOra}, Litri: {r.Litri}, Kg: {r.Kg}");
                    lista.Add(r);
                }
            }

            Debug.WriteLine($"[SGAM] Rifornimenti totali letti per i codici {string.Join(", ", codiciDaCercare)}: {lista.Count}");
            var totaleLitri = lista.Sum(r => r.Litri);
            var totaleKg = lista.Sum(r => r.Kg);
            Debug.WriteLine($"[SGAM] Somma totale litri: {totaleLitri}");
            Debug.WriteLine($"[SGAM] Somma totale kg: {totaleKg}");
            return lista;
        }

        private decimal CalcolaMediaAritmeticaGiornaliera(List<DatiConsumoRecord> dati)
        {
            if (dati == null || dati.Count == 0)
                return 0;

            // Se serve filtrare i consumi > 0
            var consumi = dati
                .Select(d => d.ConsumoKmPerLitro)
                .Where(c => c > 0)
                .ToList();

            if (consumi.Count == 0)
                return 0;

            // Semplice media dei valori di Consumo_km/l
            return consumi.Average();
        }
        private decimal CalcolaMediaSommaRifornimenti(
            List<DatiConsumoRecord> datiConsumo,
            List<RifornimentoRecord> rifornimenti,
            bool isMetano = false)
        {
            if (datiConsumo == null || datiConsumo.Count == 0)
                return 0;
            if (rifornimenti == null || rifornimenti.Count == 0)
                return 0;

            Debug.WriteLine("----- Debug dei Km giornalieri -----");
            foreach (var rec in datiConsumo)
            {
                Debug.WriteLine($"Data: {rec.Data:dd/MM/yyyy}, Km_Totali: {rec.Km_Totali}");
            }
            // Sommo tutti i km giornalieri
            decimal totKm = datiConsumo.Sum(d => d.Km_Totali);
            // Se il veicolo è a metano usiamo i Kg, altrimenti i Litri
            decimal totFuel = isMetano
                                ? rifornimenti.Sum(r => r.Kg)
                                : rifornimenti.Sum(r => r.Litri);

            Debug.WriteLine(isMetano
                ? $"Totale kg riforniti: {totFuel}"
                : $"Totale litri riforniti: {totFuel}");

            if (totKm > 0 && totFuel > 0)
                return totKm / totFuel;
            return 0;
        }


        private Dictionary<string, string> CaricaTipiCarburante()
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            using var cn = new SqlConnection("Server=srv2016app02\\sgam;Database=PARATORI;User Id=sapara;Password=S@p4ra;Encrypt=True;TrustServerCertificate=True;");

            cn.Open();

            using var cmd = new SqlCommand("Stp_AnagraficaMezziUnica", cn);
            cmd.CommandType = System.Data.CommandType.StoredProcedure;

            cmd.Parameters.AddWithValue("@Tipo", 0);
            cmd.Parameters.AddWithValue("@MostraTabella", 1);
            cmd.Parameters.AddWithValue("@Data", DateTime.Now.ToString("yyyyMMdd"));
            // ↑ se la data è fissa o la recuperi altrove, regola di conseguenza

            using var dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                var codice = dr["CODICE"] as string;               // la Targa o CodiceVeicolo
                var tipoCarb = dr["TIPO_CARBURANTE"] as string;    // "GA" o "ME"

                if (!string.IsNullOrWhiteSpace(codice) && !string.IsNullOrWhiteSpace(tipoCarb))
                {
                    result[codice.Trim()] = tipoCarb.Trim();
                }
            }

            return result;
        }
        #endregion

        #region --- AUTISTI ---

        private void CaricaDatiStessoPeriodoAutisti(System.Collections.IList autisti, DateTime da, DateTime a)
        {
            var risultati = new List<ConsumoMedioAutistaResult>();

            foreach (AutistaItem aItem in autisti)
            {
                var recs = LeggiDatiAutista(aItem.Codice, da, a);
                var media = CalcolaMediaAutista(recs);

                risultati.Add(new ConsumoMedioAutistaResult
                {
                    CodiceAutista = aItem.Codice,
                    NomeAutista = aItem.Nome,
                    MediaConsumo = media
                });
            }

            dgRisultati.ItemsSource = risultati;
        }

        private void CaricaDatiPeriodiDiversiAutisti(List<PeriodoPerVeicolo> lista)
        {
            // raggruppo tutti i record per autista
            var map = new Dictionary<string, List<DatiConsumoAutistaRecord>>();

            foreach (var per in lista)
            {
                if (!map.ContainsKey(per.Targa))
                    map[per.Targa] = new List<DatiConsumoAutistaRecord>();

                map[per.Targa].AddRange(LeggiDatiAutista(per.Targa, per.DataDa, per.DataA));
            }

            var risultati = new List<ConsumoMedioAutistaResult>();

            foreach (var k in map.Keys)
            {
                var recs = map[k];
                var media = CalcolaMediaAutista(recs);
                var nome = recs.FirstOrDefault()?.NomeAutista ?? "";

                risultati.Add(new ConsumoMedioAutistaResult
                {
                    CodiceAutista = k,
                    NomeAutista = nome,
                    MediaConsumo = media
                });
            }

            dgRisultati.ItemsSource = risultati;
        }

        private List<DatiConsumoAutistaRecord> LeggiDatiAutista(string codice, DateTime d1, DateTime d2)
        {
            var lista = new List<DatiConsumoAutistaRecord>();

            using var cn = new SqlConnection(ConnString);
            cn.Open();

            var cmd = new SqlCommand(@"
                SELECT CodiceAutista, NomeAutista, KmPercorsi, LitriConsumati,
                       KmPerLitro, DataGiorno, Veicoli
                FROM dbo.tbDatiConsumoAutisti
                WHERE CodiceAutista = @c
                  AND DataGiorno   >= @d1
                  AND DataGiorno   <= @d2
                ORDER BY DataGiorno", cn);

            cmd.Parameters.AddWithValue("@c", codice);
            cmd.Parameters.AddWithValue("@d1", d1.Date);
            cmd.Parameters.AddWithValue("@d2", d2.Date.AddDays(1).AddTicks(-1));

            using var dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                lista.Add(new DatiConsumoAutistaRecord
                {
                    CodiceAutista = dr["CodiceAutista"] as string ?? "",
                    NomeAutista = dr["NomeAutista"] as string ?? "",
                    KmPercorsi = dr["KmPercorsi"] is DBNull ? 0 : (long)dr["KmPercorsi"],
                    LitriConsumati = dr["LitriConsumati"] is DBNull ? 0m : (decimal)dr["LitriConsumati"],
                    KmPerLitro = dr["KmPerLitro"] is DBNull ? 0m : (decimal)dr["KmPerLitro"],
                    DataGiorno = dr["DataGiorno"] is DBNull ? DateTime.MinValue : (DateTime)dr["DataGiorno"],
                    Veicoli = dr["Veicoli"] as string ?? ""
                });
            }

            return lista;
        }

        private decimal CalcolaMediaAutista(List<DatiConsumoAutistaRecord> dati)
        {
            if (dati == null || dati.Count == 0)
                return 0;

            var valori = dati.Select(x => x.KmPerLitro)
                             .Where(x => x > 0)
                             .ToList();

            return valori.Count == 0 ? 0 : valori.Average();
        }

        #endregion

        #region Esporta Excel  (identico, usa dgRisultati)

        private void btnEsporta_Click(object sender, RoutedEventArgs e)
        {
            if (dgRisultati.ItemsSource == null)
                return;

            var dlg = new SaveFileDialog
            {
                Filter = "Excel file|*.xlsx",
                Title = "Salva risultati consumi"
            };
            if (dlg.ShowDialog() != true)
                return;

            using var wb = new XLWorkbook();
            var ws = wb.AddWorksheet("Consumi");

            // Scrive le intestazioni colonna
            for (int col = 0; col < dgRisultati.Columns.Count; col++)
            {
                ws.Cell(1, col + 1).Value = dgRisultati.Columns[col].Header?.ToString() ?? "";
            }

            // Scrive i dati riga per riga
            int row = 2;
            foreach (var item in dgRisultati.ItemsSource)
            {
                var props = item.GetType().GetProperties();
                for (int col = 0; col < props.Length; col++)
                {
                    // Converte sempre in stringa per evitare errori di conversione
                    var raw = props[col].GetValue(item);
                    string text = raw?.ToString() ?? "";
                    ws.Cell(row, col + 1).Value = text;
                }
                row++;
            }

            // Salva e notifica
            wb.SaveAs(dlg.FileName);
            MessageBox.Show("Esportazione completata", "OK", MessageBoxButton.OK, MessageBoxImage.Information);
        }



        #endregion
        private void tbSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            var text = tbSearch.Text.Trim();
            if (_modo == ModoRicerca.Targhe)
            {
                if (string.IsNullOrEmpty(text))
                    PopolaTarghe(_allTarghe);
                else
                    PopolaTarghe(
                      _allTarghe
                        .Where(t => t
                          .IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0));
            }
            else
            {
                if (string.IsNullOrEmpty(text))
                    PopolaAutisti(_allAutisti);
                else
                    PopolaAutisti(
                      _allAutisti
                        .Where(a => a.Codice
                          .IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0
                            || a.Nome
                          .IndexOf(text, StringComparison.OrdinalIgnoreCase) >= 0));
            }
            // resetta il pulsante Seleziona tutte
            btnToggleAll.Content = "Seleziona Tutte";
        }

        private void PopolaTarghe(IEnumerable<string> list)
        {
            _isRefreshingList = true;

            lbRisorse.Items.Clear();
            foreach (var t in list)
                lbRisorse.Items.Add(t);

            // Riallinea la selezione visibile con quella globale
            foreach (string item in lbRisorse.Items.Cast<string>())
                if (_selectedTarghe.Contains(item))
                    lbRisorse.SelectedItems.Add(item);

            _isRefreshingList = false;
        }

        private void PopolaAutisti(IEnumerable<AutistaItem> list)
        {
            _isRefreshingList = true;

            lbRisorse.Items.Clear();
            foreach (var a in list)
                lbRisorse.Items.Add(a);

            foreach (AutistaItem item in lbRisorse.Items.Cast<AutistaItem>())
                if (_selectedAutisti.Contains(item.Codice))
                    lbRisorse.SelectedItems.Add(item);

            _isRefreshingList = false;
        }


        private void LbRisorse_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (_isRefreshingList)
                return;   // IGNORA gli eventi in fase di refresh

            if (_modo == ModoRicerca.Targhe)
            {
                foreach (string added in e.AddedItems.Cast<string>())
                    _selectedTarghe.Add(added);
                foreach (string removed in e.RemovedItems.Cast<string>())
                    _selectedTarghe.Remove(removed);
            }
            else
            {
                foreach (AutistaItem added in e.AddedItems.Cast<AutistaItem>())
                    _selectedAutisti.Add(added.Codice);
                foreach (AutistaItem removed in e.RemovedItems.Cast<AutistaItem>())
                    _selectedAutisti.Remove(removed.Codice);
            }
        }


    }
}

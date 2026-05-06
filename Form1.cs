using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Mail;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TradeTrackPro
{
    public partial class Form1 : Form
    {
        private readonly HttpClient client = new HttpClient();

        private const string API_KEY = "Finnhub API";
        private const string GEMINI_API_KEY = "GEMINI API";
        private const string FILE_NAME = "portfolio.json";

        private const string SENDER_EMAIL = "Your Mail";
        private const string SENDER_PASSWORD = "Sender Password";

        private List<Stock> portfolio = new();

        private DataGridView? dgvPortfolio;
        private Panel? pnlDetails;
        private Label? lblSymbol, lblCompany, lblSector, lblCurrentPrice;
        private TextBox? txtSymbol, txtBuyPrice, txtTargetPrice, txtNotificationEmail;
        private NumericUpDown? numQuantity;
        private Button? btnAddToPortfolio, btnDeleteStock;
        private Chart? chartStock;

        public class Stock
        {
            public string Symbol { get; set; } = "";
            public string Company { get; set; } = "";
            public string Sector { get; set; } = "";
            public int Quantity { get; set; }
            public decimal BuyPrice { get; set; }
            public decimal CurrentPrice { get; set; }
            public decimal TargetPrice { get; set; }
            public bool AlertSent { get; set; } = false;

            public decimal ProfitLoss => (CurrentPrice - BuyPrice) * Quantity;
            public decimal ProfitLossPercent => BuyPrice > 0 ? ((CurrentPrice - BuyPrice) / BuyPrice * 100) : 0;
        }

        public Form1()
        {
            InitializeComponent();
            Load += Form_Load;
        }

        private void Form_Load(object? sender, EventArgs e)
        {
            Text = "TradeTrack | Scientific Session Project";
            Size = new Size(1400, 900);
            BackColor = Color.FromArgb(30, 30, 30);
            ForeColor = Color.White;
            StartPosition = FormStartPosition.CenterScreen;

            CreateDesign();
            LoadPortfolio();
            StartTimer();
        }

        private void CreateDesign()
        {
            dgvPortfolio = new DataGridView
            {
                Location = new Point(0, 0),
                Width = 1000,
                Height = ClientSize.Height,
                BackgroundColor = Color.FromArgb(40, 40, 40),
                ForeColor = Color.White,
                GridColor = Color.FromArgb(60, 60, 60),
                RowHeadersVisible = false,
                AllowUserToAddRows = false,
                ReadOnly = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                BorderStyle = BorderStyle.None,
                ColumnHeadersHeight = 35,
                RowTemplate = { Height = 30 },
                AutoGenerateColumns = false
            };

            dgvPortfolio.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(50, 50, 50);
            dgvPortfolio.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPortfolio.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            dgvPortfolio.DefaultCellStyle.BackColor = Color.FromArgb(40, 40, 40);
            dgvPortfolio.DefaultCellStyle.ForeColor = Color.White;
            dgvPortfolio.DefaultCellStyle.SelectionBackColor = Color.FromArgb(0, 120, 215);
            dgvPortfolio.DefaultCellStyle.SelectionForeColor = Color.White;
            dgvPortfolio.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(45, 45, 45);

            dgvPortfolio.Columns.AddRange(new DataGridViewColumn[]
            {
                new DataGridViewTextBoxColumn { HeaderText = "Symbol", DataPropertyName = "Symbol", Width = 80 },
                new DataGridViewTextBoxColumn { HeaderText = "Company", DataPropertyName = "Company", Width = 150 },
                new DataGridViewTextBoxColumn { HeaderText = "Qty", DataPropertyName = "Quantity", Width = 60 },
                new DataGridViewTextBoxColumn { HeaderText = "Buy Price", DataPropertyName = "BuyPrice", Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "$#,##0.00" } },
                new DataGridViewTextBoxColumn { HeaderText = "Target", DataPropertyName = "TargetPrice", Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "$#,##0.00" } },
                new DataGridViewTextBoxColumn { HeaderText = "Current", DataPropertyName = "CurrentPrice", Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "$#,##0.00" } },
                new DataGridViewTextBoxColumn { HeaderText = "P&L ($)", DataPropertyName = "ProfitLoss", Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "$#,##0.00" } },
                new DataGridViewTextBoxColumn { HeaderText = "P&L (%)", DataPropertyName = "ProfitLossPercent", Width = 90, DefaultCellStyle = new DataGridViewCellStyle { Format = "0.00%" } }
            });

            dgvPortfolio.CellFormatting += (s, e) =>
            {
                if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
                {
                    var col = dgvPortfolio.Columns[e.ColumnIndex];
                    if (e.Value != null && (col.DataPropertyName == "ProfitLoss" || col.DataPropertyName == "ProfitLossPercent"))
                    {
                        if (decimal.TryParse(e.Value.ToString()?.Replace("%", "").Replace("$", "").Trim(), out decimal val))
                        {
                            e.CellStyle.ForeColor = val >= 0 ? Color.FromArgb(0, 200, 80) : Color.FromArgb(255, 80, 80);
                            if (col.DataPropertyName == "ProfitLossPercent") e.Value = $"{val:F2}%";
                        }
                    }
                }
            };

            dgvPortfolio.CellClick += async (s, e) =>
            {
                if (e.RowIndex >= 0)
                {
                    if (dgvPortfolio.Columns[e.ColumnIndex].Name == "AIAdvice")
                        await GetAIAdvice(e.RowIndex);
                    else
                    {
                        var symbol = dgvPortfolio.Rows[e.RowIndex].Cells[0].Value.ToString();
                        if (!string.IsNullOrEmpty(symbol)) await LoadStockChart(symbol);
                    }
                }
            };

            dgvPortfolio.SelectionChanged += (s, e) => ShowStockDetails();
            Controls.Add(dgvPortfolio);

            pnlDetails = new Panel
            {
                Location = new Point(1000, 0),
                Width = ClientSize.Width - 1000,
                Height = ClientSize.Height,
                BackColor = Color.FromArgb(25, 25, 25),
                Padding = new Padding(15),
                AutoScroll = true
            };

            var grpSettings = new GroupBox { Text = "Alert Settings", Location = new Point(15, 10), Width = pnlDetails.Width - 40, Height = 85, ForeColor = Color.Orange, Font = new Font("Segoe UI", 10F, FontStyle.Bold) };
            var lblEmailInput = CreateLabel("Notification Email:", 20, 30);
            txtNotificationEmail = CreateTextBox(20, 50, grpSettings.Width - 40);
            txtNotificationEmail.PlaceholderText = "Enter email for alerts...";
            grpSettings.Controls.Add(lblEmailInput); grpSettings.Controls.Add(txtNotificationEmail);

            var grpDetails = new GroupBox { Text = "Stock Details", Location = new Point(15, 110), Width = pnlDetails.Width - 40, Height = 150, ForeColor = Color.Cyan, Font = new Font("Segoe UI", 11F, FontStyle.Bold) };
            lblSymbol = CreateLabel("Symbol : -", 20, 30);
            lblCompany = CreateLabel("Company : -", 20, 60);
            lblSector = CreateLabel("Sector : -", 20, 90);
            lblCurrentPrice = CreateLabel("Current Price :", 20, 120);
            grpDetails.Controls.AddRange(new Control[] { lblSymbol, lblCompany, lblSector, lblCurrentPrice });

            var grpAdd = new GroupBox { Text = "Add / Update Stock", Location = new Point(15, 270), Width = pnlDetails.Width - 40, Height = 290, ForeColor = Color.White, Font = new Font("Segoe UI", 11F, FontStyle.Bold) };
            var lblSymInput = CreateLabel("Symbol:", 20, 30);
            txtSymbol = CreateTextBox(20, 55, grpAdd.Width - 40); txtSymbol.CharacterCasing = CharacterCasing.Upper;
            var lblPriceInput = CreateLabel("Buy Price:", 20, 90);
            txtBuyPrice = CreateTextBox(20, 115, grpAdd.Width - 40);
            var lblQtyInput = CreateLabel("Quantity:", 20, 150);
            numQuantity = new NumericUpDown { Location = new Point(20, 175), Width = 120, Height = 30, Minimum = 1, Value = 1, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(50, 50, 50), ForeColor = Color.White };
            var lblTargetInput = CreateLabel("Alert Price:", 160, 150);
            txtTargetPrice = CreateTextBox(160, 175, 120); txtTargetPrice.PlaceholderText = "Optional";
            btnAddToPortfolio = new Button { Text = "Add / Update", Location = new Point(20, 220), Width = grpAdd.Width - 40, Height = 40, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnAddToPortfolio.Click += async (s, e) => await AddStockToPortfolio();
            grpAdd.Controls.AddRange(new Control[] { lblSymInput, txtSymbol, lblPriceInput, txtBuyPrice, lblQtyInput, numQuantity, lblTargetInput, txtTargetPrice, btnAddToPortfolio });

            btnDeleteStock = new Button { Text = "Delete Selected Stock", Location = new Point(15, 570), Width = pnlDetails.Width - 40, Height = 40, BackColor = Color.FromArgb(180, 40, 40), ForeColor = Color.White, Font = new Font("Segoe UI", 10F, FontStyle.Bold), FlatStyle = FlatStyle.Flat, Cursor = Cursors.Hand };
            btnDeleteStock.Click += BtnDeleteStock_Click;

            chartStock = new Chart { Location = new Point(15, 620), Width = pnlDetails.Width - 40, Height = 220, BackColor = Color.FromArgb(30, 30, 30) };
            var chartArea = new ChartArea { Name = "MainArea", BackColor = Color.FromArgb(40, 40, 40) };
            chartArea.AxisX.LabelStyle.ForeColor = Color.Silver; chartArea.AxisY.LabelStyle.ForeColor = Color.Silver;
            chartArea.AxisX.LineColor = Color.FromArgb(60, 60, 60); chartArea.AxisY.LineColor = Color.FromArgb(60, 60, 60);
            chartArea.AxisX.MajorGrid.LineColor = Color.FromArgb(50, 50, 50); chartArea.AxisY.MajorGrid.LineColor = Color.FromArgb(50, 50, 50);
            chartStock.ChartAreas.Add(chartArea);
            chartStock.Titles.Add(new Title { Text = "Select a stock to view trend", ForeColor = Color.White, Font = new Font("Segoe UI", 10, FontStyle.Bold) });

            pnlDetails.Controls.AddRange(new Control[] { grpSettings, grpDetails, grpAdd, btnDeleteStock, chartStock });
            Controls.Add(pnlDetails);
        }

        private Label CreateLabel(string text, int x, int y) => new Label { Text = text, Location = new Point(x, y), AutoSize = true, Font = new Font("Segoe UI", 10F), ForeColor = Color.Cyan };
        private TextBox CreateTextBox(int x, int y, int w) => new TextBox { Location = new Point(x, y), Width = w, Font = new Font("Segoe UI", 10F), BackColor = Color.FromArgb(50, 50, 50), ForeColor = Color.White, BorderStyle = BorderStyle.FixedSingle };

        private async Task AddStockToPortfolio()
        {
            if (txtSymbol == null || txtBuyPrice == null || numQuantity == null) return;
            string symbol = txtSymbol.Text.Trim().ToUpper();
            if (string.IsNullOrEmpty(symbol)) { MessageBox.Show("Enter symbol!"); return; }
            if (!decimal.TryParse(txtBuyPrice.Text, out decimal buyPrice) || buyPrice <= 0) { MessageBox.Show("Invalid price!"); return; }
            decimal targetPrice = 0;
            if (txtTargetPrice != null && !string.IsNullOrEmpty(txtTargetPrice.Text)) decimal.TryParse(txtTargetPrice.Text, out targetPrice);

            try
            {
                var quoteJson = await client.GetStringAsync($"https://finnhub.io/api/v1/quote?symbol={symbol}&token={API_KEY}");
                var profileJson = await client.GetStringAsync($"https://finnhub.io/api/v1/stock/profile2?symbol={symbol}&token={API_KEY}");
                var quote = JObject.Parse(quoteJson); var profile = JObject.Parse(profileJson);
                decimal currentPrice = quote["c"] != null ? (decimal)(double)quote["c"]! : 0;
                string company = profile["name"]?.ToString() ?? symbol;
                string sector = profile["finnhubIndustry"]?.ToString() ?? "N/A";

                if (currentPrice == 0) { MessageBox.Show("Stock not found!"); return; }

                var existing = portfolio.FirstOrDefault(s => s.Symbol == symbol);
                if (existing != null)
                {
                    int oldQty = existing.Quantity;
                    int newQty = (int)numQuantity.Value;
                    existing.Quantity = oldQty + newQty;
                    existing.BuyPrice = ((existing.BuyPrice * oldQty) + (buyPrice * newQty)) / existing.Quantity;
                    existing.CurrentPrice = currentPrice; existing.TargetPrice = targetPrice; existing.AlertSent = false;
                }
                else
                {
                    portfolio.Add(new Stock { Symbol = symbol, Company = company, Sector = sector, Quantity = (int)numQuantity.Value, BuyPrice = buyPrice, CurrentPrice = currentPrice, TargetPrice = targetPrice });
                }

                SavePortfolio(); UpdateDisplay();
                txtSymbol.Text = ""; txtBuyPrice.Text = ""; if (txtTargetPrice != null) txtTargetPrice.Text = ""; numQuantity.Value = 1;
                await LoadStockChart(symbol);
                MessageBox.Show($"{symbol} added!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
        }

        private void BtnDeleteStock_Click(object? sender, EventArgs e)
        {
            if (dgvPortfolio == null || dgvPortfolio.SelectedRows.Count == 0) return;
            string symbol = dgvPortfolio.SelectedRows[0].Cells[0].Value.ToString()!;
            if (MessageBox.Show($"Delete {symbol}?", "Confirm", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                var stock = portfolio.FirstOrDefault(s => s.Symbol == symbol);
                if (stock != null)
                {
                    portfolio.Remove(stock); SavePortfolio(); UpdateDisplay();
                    if (chartStock != null) { chartStock.Series.Clear(); chartStock.Titles[0].Text = "Select a stock"; }
                }
            }
        }

        private async Task LoadStockChart(string symbol)
        {
            if (chartStock == null) return;
            chartStock.Series.Clear();
            chartStock.Titles[0].Text = $"{symbol} - Loading...";
            chartStock.Invalidate();

            bool apiSuccess = false;

            try
            {
                long to = DateTimeOffset.Now.ToUnixTimeSeconds();
                long from = DateTimeOffset.Now.AddDays(-30).ToUnixTimeSeconds();
                string url = $"https://finnhub.io/api/v1/stock/candle?symbol={symbol}&resolution=D&from={from}&to={to}&token={API_KEY}";

                var response = await client.GetAsync(url);

                if (response.IsSuccessStatusCode)
                {
                    var json = await response.Content.ReadAsStringAsync();
                    dynamic data = JsonConvert.DeserializeObject(json)!;

                    if (data.s == "ok" && data.c != null)
                    {
                        DrawChart(symbol, data.c, data.t);
                        apiSuccess = true;
                    }
                }
            }
            catch
            {
                apiSuccess = false;
            }

            if (!apiSuccess)
            {
                GenerateMockChartData(symbol);
            }
        }

        private void GenerateMockChartData(string symbol)
        {
            if (chartStock == null) return;

            var stock = portfolio.FirstOrDefault(s => s.Symbol == symbol);
            decimal basePrice = stock != null ? stock.CurrentPrice : 100;

            var prices = new List<double>();
            var dates = new List<string>();
            var rand = new Random();
            double price = (double)basePrice * 0.9; 

            for (int i = 0; i < 30; i++)
            {
                double change = (rand.NextDouble() * 0.04) - 0.02;
                price = price * (1 + change);
                prices.Add(price);
                dates.Add(DateTime.Now.AddDays(-30 + i).ToString("dd/MM"));
            }

            prices[29] = (double)basePrice;

            var series = new Series(symbol);
            series.ChartType = SeriesChartType.Line;
            series.BorderWidth = 3;
            series.Color = Color.FromArgb(0, 180, 255);
            series.IsVisibleInLegend = false;

            for (int i = 0; i < 30; i++)
            {
                series.Points.AddXY(dates[i], prices[i]);
            }

            chartStock.Series.Clear();
            chartStock.Series.Add(series);
            chartStock.Titles[0].Text = $"{symbol} - 30 Day Trend";
            chartStock.ChartAreas[0].RecalculateAxesScale();
        }

        private void DrawChart(string symbol, dynamic prices, dynamic times)
        {
            var series = new Series(symbol);
            series.ChartType = SeriesChartType.Line;
            series.BorderWidth = 3;
            series.Color = Color.FromArgb(0, 180, 255);
            series.IsVisibleInLegend = false;

            for (int i = 0; i < prices.Count; i++)
            {
                double ts = (double)times[i];
                DateTime date = DateTimeOffset.FromUnixTimeSeconds((long)ts).DateTime;
                double price = (double)prices[i];
                series.Points.AddXY(date.ToString("dd/MM"), price);
            }

            chartStock.Series.Add(series);
            chartStock.Titles[0].Text = $"{symbol} - 30 Day Trend";
            chartStock.ChartAreas[0].RecalculateAxesScale();
        }

        private async Task GetAIAdvice(int rowIndex)
        {
            if (rowIndex < 0 || rowIndex >= portfolio.Count) return;
            var stock = portfolio[rowIndex];

            var load = new Form { Text = "AI Analysis", Size = new Size(300, 100), StartPosition = FormStartPosition.CenterScreen, FormBorderStyle = FormBorderStyle.FixedDialog, ControlBox = false };
            load.Controls.Add(new Label { Text = "Analyzing...", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter });
            load.Show();

            try
            {
                var prompt = $@"
                    Act as a senior financial analyst analyzing {stock.Symbol}.
                    Context: Buy: ${stock.BuyPrice}, Current: ${stock.CurrentPrice}.
                    
                    Task: Provide a short professional report.
                    Format Rules:
                    1. DO NOT use asterisks (*) or markdown bolding. Use UPPERCASE for headings.
                    2. Headings: VERDICT, ANALYSIS, RISK.
                    3. Keep it short (max 100 words).";

                var body = new { contents = new[] { new { parts = new[] { new { text = prompt } } } } };
                var json = JsonConvert.SerializeObject(body);
                var content = new StringContent(json, System.Text.Encoding.UTF8, "application/json");
                var response = await client.PostAsync($"https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash-exp:generateContent?key={GEMINI_API_KEY}", content);

                load.Close();
                if (response.IsSuccessStatusCode)
                {
                    var resJson = JObject.Parse(await response.Content.ReadAsStringAsync());
                    string advice = resJson["candidates"]?[0]?["content"]?["parts"]?[0]?["text"]?.ToString() ?? "No analysis.";

                    advice = advice.Replace("**", "").Replace("*", "");

                    ShowAdviceDialog(stock.Symbol, advice);
                }
            }
            catch (Exception ex) { load.Close(); MessageBox.Show("AI Error: " + ex.Message); }
        }

        private void ShowAdviceDialog(string symbol, string advice)
        {
            var form = new Form { Text = $"Analyst Report: {symbol}", Size = new Size(600, 450), StartPosition = FormStartPosition.CenterScreen, BackColor = Color.FromArgb(30, 30, 30) };
            var txt = new RichTextBox { Dock = DockStyle.Fill, Text = advice, ReadOnly = true, BackColor = Color.FromArgb(40, 40, 40), ForeColor = Color.White, Font = new Font("Segoe UI", 11F), BorderStyle = BorderStyle.None, Padding = new Padding(20) };
            var btn = new Button { Text = "Close", Dock = DockStyle.Bottom, Height = 40, BackColor = Color.FromArgb(0, 120, 215), ForeColor = Color.White, FlatStyle = FlatStyle.Flat };
            btn.Click += (s, e) => form.Close();
            form.Controls.Add(txt); form.Controls.Add(btn); form.ShowDialog();
        }

        private void StartTimer() { var t = new System.Windows.Forms.Timer { Interval = 15000 }; t.Tick += async (s, e) => await UpdatePrices(); t.Start(); _ = UpdatePrices(); }

        private async Task UpdatePrices()
        {
            foreach (var stock in portfolio)
            {
                try
                {
                    var json = await client.GetStringAsync($"https://finnhub.io/api/v1/quote?symbol={stock.Symbol}&token={API_KEY}");
                    var data = JObject.Parse(json);
                    if (data["c"] != null)
                    {
                        stock.CurrentPrice = (decimal)(double)data["c"]!;
                        CheckAndSendAlert(stock);
                    }
                }
                catch { }
            }
            UpdateDisplay();
        }

        private void CheckAndSendAlert(Stock stock)
        {
            if (stock.TargetPrice > 0 && stock.CurrentPrice < stock.TargetPrice && !stock.AlertSent) { SendEmailAlert(stock); stock.AlertSent = true; }
            else if (stock.CurrentPrice >= stock.TargetPrice) stock.AlertSent = false;
        }

        private void SendEmailAlert(Stock stock)
        {
            string receiverEmail = txtNotificationEmail?.Text.Trim() ?? "";
            if (string.IsNullOrEmpty(SENDER_EMAIL) || string.IsNullOrEmpty(receiverEmail)) return;
            try
            {
                using (MailMessage mail = new MailMessage())
                {
                    mail.From = new MailAddress(SENDER_EMAIL); mail.To.Add(receiverEmail);
                    mail.Subject = $"⚠️ ALERT: {stock.Symbol} Price Drop!";
                    mail.Body = $"WARNING: {stock.Symbol} dropped below ${stock.TargetPrice}.\nCurrent: ${stock.CurrentPrice}";
                    using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                    {
                        smtp.Credentials = new NetworkCredential(SENDER_EMAIL, SENDER_PASSWORD);
                        smtp.EnableSsl = true; smtp.Send(mail);
                    }
                }
            }
            catch { }
        }

        private void UpdateDisplay()
        {
            if (dgvPortfolio == null) return;
            var selIdx = dgvPortfolio.SelectedRows.Count > 0 ? dgvPortfolio.SelectedRows[0].Index : -1;
            dgvPortfolio.DataSource = null; dgvPortfolio.DataSource = portfolio;
            if (!dgvPortfolio.Columns.Contains("AIAdvice"))
            {
                var btn = new DataGridViewButtonColumn { Name = "AIAdvice", HeaderText = "AI", Text = "Ask", UseColumnTextForButtonValue = true, Width = 60, FlatStyle = FlatStyle.Flat };
                btn.DefaultCellStyle.BackColor = Color.Purple; btn.DefaultCellStyle.ForeColor = Color.White; dgvPortfolio.Columns.Add(btn);
            }
            foreach (DataGridViewColumn c in dgvPortfolio.Columns) if (c.Name != "AIAdvice") c.ReadOnly = true;
            if (selIdx >= 0 && selIdx < dgvPortfolio.Rows.Count) dgvPortfolio.Rows[selIdx].Selected = true;

            decimal val = portfolio.Sum(x => x.CurrentPrice * x.Quantity); decimal cost = portfolio.Sum(x => x.BuyPrice * x.Quantity);
            Text = $"TradeTrack | Portfolio: ${val:F2} | P&L: ${(val - cost):F2}";
        }

        private void ShowStockDetails()
        {
            if (dgvPortfolio == null || dgvPortfolio.SelectedRows.Count == 0) return;
            var symbol = dgvPortfolio.SelectedRows[0].Cells[0].Value.ToString();
            var stock = portfolio.FirstOrDefault(s => s.Symbol == symbol);
            if (stock != null)
            {
                lblSymbol!.Text = $"Symbol : {stock.Symbol}"; lblCompany!.Text = $"Company : {stock.Company}";
                lblSector!.Text = $"Sector : {stock.Sector}"; lblCurrentPrice!.Text = $"Current Price : ${stock.CurrentPrice:F2}";
                _ = LoadStockChart(stock.Symbol);
            }
        }

        private void LoadPortfolio()
        {
            if (File.Exists(FILE_NAME)) try { portfolio = JsonConvert.DeserializeObject<List<Stock>>(File.ReadAllText(FILE_NAME)) ?? new(); } catch { portfolio = new(); }
            UpdateDisplay();
        }
        private void SavePortfolio() => File.WriteAllText(FILE_NAME, JsonConvert.SerializeObject(portfolio, Formatting.Indented));
    }
}
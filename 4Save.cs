using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using Dapper;
using HtmlAgilityPack;

namespace _4Save
{
    public class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        private TextBox txtFolderPath;
        private Button btnBrowse;
        private Button btnScan;
        private ListView listResults;
        private Label lblStatus;
        private string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "db.sqlite");

        public MainForm()
        {
            InitializeComponent();
            InitializeDatabase();
        }

        private void InitializeComponent()
        {
            this.Text = "4Save (CUSA ID Title Lookup)";
            this.Size = new System.Drawing.Size(1000, 600);
            this.MinimumSize = new System.Drawing.Size(800, 400);

            // Folder path selection
            Label lblFolderPath = new()

            {
                Text = "Select Source Folder:",
                Location = new System.Drawing.Point(10, 15),
                AutoSize = true
            };

            txtFolderPath = new TextBox
            {
                Location = new System.Drawing.Point(150, 12),
                Width = 500
            };

            btnBrowse = new Button
            {
                Text = "Browse",
                Location = new System.Drawing.Point(660, 10),
                Width = 80
            };
            btnBrowse.Click += BtnBrowse_Click;

            btnScan = new Button
            {
                Text = "Scan",
                Location = new System.Drawing.Point(660, 45),
                Width = 80
            };
            btnScan.Click += BtnScan_Click;

            // Status label
            lblStatus = new Label
            {
                Text = "Ready",
                Location = new System.Drawing.Point(10, 50),
                AutoSize = true
            };

            // Results list
            listResults = new ListView
            {
                Location = new System.Drawing.Point(10, 80),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                Width = this.ClientSize.Width - 20,
                Height = this.ClientSize.Height - 90,
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Sorting = SortOrder.Ascending
            };

            listResults.Columns.Add("CUSA ID", 100);
            listResults.Columns.Add("Title", 300);
            listResults.Columns.Add("Date", 180);
            listResults.Columns.Add("Actions", 150);

            // Enable double buffering for smoother UI
            typeof(ListView).InvokeMember("DoubleBuffered",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
                null, listResults, new object[] { true });

            // Setup ListViewSubItem click handler
            listResults.MouseClick += ListResults_MouseClick;

            // Add controls to form
            this.Controls.Add(lblFolderPath);
            this.Controls.Add(txtFolderPath);
            this.Controls.Add(btnBrowse);
            this.Controls.Add(btnScan);
            this.Controls.Add(lblStatus);
            this.Controls.Add(listResults);
        }

        private void InitializeDatabase()
        {
            try
            {
                if (!File.Exists(dbPath))
                {
                    SQLiteConnection.CreateFile(dbPath);
                    using var connection = GetConnection();
                    connection.Open();
                    connection.Execute(
                        @"CREATE TABLE IF NOT EXISTS CusaTitles (
                                CusaId TEXT PRIMARY KEY,
                                Title TEXT NOT NULL,
                                LastUpdated TEXT NOT NULL
                            )");
                }
                // Remove the schema update code for ImageUrl since we don't need it anymore
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error initializing database: {ex.Message}", "Database Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private SQLiteConnection GetConnection()
        {
            return new SQLiteConnection($"Data Source={dbPath};Version=3;");
        }

        private void BtnBrowse_Click(object? sender, EventArgs e)
        {
            using FolderBrowserDialog folderDialog = new();
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                txtFolderPath.Text = folderDialog.SelectedPath;

            }
        }

        private async void BtnScan_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtFolderPath.Text) || !Directory.Exists(txtFolderPath.Text))
            {
                MessageBox.Show("Please select a valid folder path.", "Invalid Path", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnScan.Enabled = false;
            listResults.Items.Clear();
            lblStatus.Text = "Scanning...";

            try
            {
                string[] directories = Directory.GetDirectories(txtFolderPath.Text);
                List<CusaInfo> cusaInfoList = new();

                // Extract CUSA IDs from folder names and get dates
                foreach (string dir in directories)
                {
                    string dirName = Path.GetFileName(dir);
                    Match match = Regex.Match(dirName, @"CUSA\d{5}", RegexOptions.IgnoreCase);
                    if (match.Success)
                    {
                        cusaInfoList.Add(new CusaInfo
                        {
                            CusaId = match.Value.ToUpper(),
                            Directory = dir,
                            Date = GetFolderDate(dir)
                        });
                    }
                }

                // Update status
                lblStatus.Text = $"Found {cusaInfoList.Count} CUSA IDs. Looking up titles...";

                // First check the database for cached titles
                using (var connection = GetConnection())
                {
                    connection.Open();
                    var cachedInfo = connection.Query<CusaTitle>(
                        "SELECT CusaId, Title FROM CusaTitles WHERE CusaId IN @CusaIds",
                        new { CusaIds = cusaInfoList.Select(c => c.CusaId).ToArray() }
                    ).ToDictionary(c => c.CusaId);

                    // Add cached info to the list
                    foreach (var info in cusaInfoList.Where(c => cachedInfo.ContainsKey(c.CusaId)))
                    {
                        info.Title = cachedInfo[info.CusaId].Title;
                    }
                }

                // Lookup titles for items not in the database
                int processed = 0;
                int totalToProcess = cusaInfoList.Count(c => string.IsNullOrEmpty(c.Title));
                int totalItems = cusaInfoList.Count;

                foreach (CusaInfo info in cusaInfoList.Where(c => string.IsNullOrEmpty(c.Title)))
                {
                    processed++;
                    lblStatus.Text = $"Processing: {processed}/{totalToProcess} (Total: {totalItems})";
                    Application.DoEvents();

                    await LookupAndSaveTitle(info);
                }

                // Display all results
                foreach (CusaInfo info in cusaInfoList.OrderBy(c => c.CusaId))
                {
                    ListViewItem item = new(info.CusaId);

                    // Make the title item look like a link if we have an image
                    string titleText = info.Title ?? "Not found";
                    item.SubItems.Add(titleText);

                    item.SubItems.Add(info.Date?.ToString("yyyy-MM-dd HH:mm:ss") ?? "Unknown");

                    // Add actions column with Open and Delete buttons
                    var actionsSubItem = item.SubItems.Add("📁 Open | 🗑️ Delete");

                    // Store the folder path in the item's tag
                    item.Tag = info;

                    listResults.Items.Add(item);
                }

                lblStatus.Text = $"Completed. Found {cusaInfoList.Count} CUSA IDs.";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                lblStatus.Text = "Error occurred during scanning";
            }
            finally
            {
                btnScan.Enabled = true;
            }
        }

        private async Task LookupAndSaveTitle(CusaInfo info)
        {
            // First try orbispatches.com
            string? title = null;

            var orbisResult = await GetInfoFromOrbisPatches(info.CusaId);
            if (!string.IsNullOrEmpty(orbisResult))
            {
                title = orbisResult;
            }

            // If not found, try serialstation.com
            if (string.IsNullOrEmpty(title))
            {
                var serialResult = await GetInfoFromSerialStation(info.CusaId);
                if (!string.IsNullOrEmpty(serialResult))
                {
                    title = serialResult;
                }
            }

            info.Title = title ?? string.Empty;

            // Save to database if title was found
            if (!string.IsNullOrEmpty(title))
            {
                try
                {
                    using (var connection = GetConnection())
                    {
                        connection.Open();
                        connection.Execute(
                            @"INSERT OR REPLACE INTO CusaTitles (CusaId, Title, LastUpdated) 
                              VALUES (@CusaId, @Title, @LastUpdated)",
                            new
                            {
                                CusaId = info.CusaId,
                                Title = info.Title,
                                LastUpdated = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                            });
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving to database: {ex.Message}");
                }
            }
        }

        private async Task<string?> GetInfoFromOrbisPatches(string cusaId)
        {
            try
            {
                string url = $"https://orbispatches.com/{cusaId}";
                HtmlWeb web = new();
                web.AutoDetectEncoding = true;
                web.OverrideEncoding = Encoding.UTF8;
                HtmlAgilityPack.HtmlDocument doc = await Task.Run(() => web.Load(url));

                string title = null;

                // Parse title
                HtmlNode titleNode = doc.DocumentNode.SelectSingleNode("//header//h1[@class='bd-title']");
                if (titleNode != null)
                {
                    title = HttpUtility.HtmlDecode(titleNode.InnerText.Trim());
                    title = CleanupTitle(title);
                }

                return title ?? null;
            }
            catch
            {
                return null;
            }
        }

        private async Task<string> GetInfoFromSerialStation(string cusaId)
        {
            try
            {
                // Split CUSA ID for serialstation format
                string mainPart = cusaId.Substring(0, 4); // "CUSA"
                string numberPart = cusaId.Substring(4);  // "00031"

                string url = $"https://serialstation.com/titles/{mainPart}/{numberPart}";
                HtmlWeb web = new();
                web.AutoDetectEncoding = true;
                web.OverrideEncoding = Encoding.UTF8;
                HtmlAgilityPack.HtmlDocument doc = await Task.Run(() => web.Load(url));

                string title = null;

                // Parse title
                HtmlNode titleNode = doc.DocumentNode.SelectSingleNode("//main[contains(@class,'container')]//h1");
                if (titleNode != null)
                {
                    title = HttpUtility.HtmlDecode(titleNode.InnerText.Trim());
                    title = CleanupTitle(title);
                }

                return title;
            }
            catch
            {
                return null;
            }
        }

        private string CleanupTitle(string title)
        {
            if (string.IsNullOrEmpty(title))
                return title;

            // Remove any trailing unnecessary text that might appear in titles
            title = Regex.Replace(title, @"\s*\(CUSA\d+\)\s*$", "", RegexOptions.IgnoreCase);

            // Remove any duplicate spaces
            title = Regex.Replace(title, @"\s+", " ");

            return title.Trim();
        }

        private DateTime? GetFolderDate(string folderPath)
        {
            try
            {
                // Try to find any files in the directory
                var files = Directory.GetFiles(folderPath, "*", SearchOption.AllDirectories);

                if (files.Length > 0)
                {
                    // Get the earliest creation date from all files
                    return files.Select(f => File.GetCreationTime(f))
                                .OrderBy(date => date)
                                .FirstOrDefault();
                }

                // If no files, use the folder creation date
                return Directory.GetCreationTime(folderPath);
            }
            catch
            {
                return null;
            }
        }

        private void ListResults_MouseClick(object? sender, MouseEventArgs e)
        {
            // Get the clicked item
            ListViewItem item = listResults.GetItemAt(e.X, e.Y);
            if (item == null) return;

            // Get the subitem (column) that was clicked
            ListViewHitTestInfo hitTest = listResults.HitTest(e.X, e.Y);
            if (hitTest.SubItem == null) return;

            int columnIndex = hitTest.Item.SubItems.IndexOf(hitTest.SubItem);

            if (item.Tag is CusaInfo info)
            {
                // Only handle Actions column (column index 3)
                if (columnIndex == 3)
                {
                    // Determine if Open or Delete was clicked based on X position
                    string actionText = hitTest.SubItem.Text;
                    int separatorIndex = actionText.IndexOf("|");
                    int openWidth = TextRenderer.MeasureText(actionText.Substring(0, separatorIndex), listResults.Font).Width;

                    // Open folder action (left side of the separator)
                    if (e.X - hitTest.SubItem.Bounds.Left <= openWidth)
                    {
                        if (Directory.Exists(info.Directory))
                        {
                            try
                            {
                                Process.Start("explorer.exe", info.Directory);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"Error opening folder: {ex.Message}", "Error",
                                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            MessageBox.Show("Folder does not exist.", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                    // Delete action (right side of the separator)
                    else
                    {
                        DeleteCusaFromDatabase(info.CusaId, item);
                    }
                }
            }
        }

        private void DeleteCusaFromDatabase(string cusaId, ListViewItem item)
        {
            // Confirm deletion with user
            var result = MessageBox.Show($"Are you sure you want to delete {cusaId} from the database?",
                "Confirm Deletion", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                try
                {
                    using (var connection = GetConnection())
                    {
                        connection.Open();
                        int affected = connection.Execute(
                            "DELETE FROM CusaTitles WHERE CusaId = @CusaId",
                            new { CusaId = cusaId });

                        if (affected > 0)
                        {
                            // Remove from ListView
                            listResults.Items.Remove(item);
                            lblStatus.Text = $"Deleted {cusaId} from database.";
                        }
                        else
                        {
                            // Item might not be in DB but still in list
                            MessageBox.Show($"Item {cusaId} was not found in the database.",
                                "Not Found", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error deleting item: {ex.Message}", "Database Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }
    }

    public class CusaInfo
    {
        public string CusaId { get; set; }
        public string Title { get; set; }
        public string Directory { get; set; }
        public DateTime? Date { get; set; }
    }

    public class CusaTitle
    {
        public string CusaId { get; set; }
        public string Title { get; set; }
        public string LastUpdated { get; set; }
    }
}
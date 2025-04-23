using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;
using Dapper;
using HtmlAgilityPack;

[assembly: System.Reflection.AssemblyMetadata("SQLite", "System.Data.SQLite")]

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

            // Disable scan button initially
            btnScan.Enabled = false;
            txtFolderPath.TextChanged += TxtFolderPath_TextChanged;
        }

        // Toggle scan button based on folder path validity
        private void TxtFolderPath_TextChanged(object? sender, EventArgs e)
        {
            btnScan.Enabled = !string.IsNullOrWhiteSpace(txtFolderPath.Text) && Directory.Exists(txtFolderPath.Text);
        }

        private void InitializeComponent()
        {
            this.Text = "4Save (PS4 & PS5 Game ID Lookup Tool)";
            this.Size = new System.Drawing.Size(850, 600);
            this.MinimumSize = new System.Drawing.Size(850, 400);
            this.MaximumSize = new System.Drawing.Size(850, 2000);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            // Set the form icon
            string iconPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "icon.ico");
            if (File.Exists(iconPath))
            {
                try
                {
                    this.Icon = new Icon(iconPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error loading icon: {ex.Message}");
                }
            }

            // Folder path selection
            Label lblFolderPath = new()

            {
                Text = "Source Folder:",
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
                Width = 100
            };
            btnBrowse.Click += BtnBrowse_Click;

            btnScan = new Button
            {
                Text = "Scan",
                Location = new System.Drawing.Point(660, 45),
                Width = 100
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

            listResults.Columns.Add("ID", 100);
            listResults.Columns.Add("Title", 250);
            listResults.Columns.Add("Platform", 60);
            listResults.Columns.Add("Date", 180);
            listResults.Columns.Add("Actions", 200); // Increased from 150 to 200

            // Enable double buffering for smoother UI
            typeof(ListView).InvokeMember("DoubleBuffered",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.SetProperty,
                null, listResults, new object[] { true });

            // Setup ListViewSubItem click handler
            listResults.MouseClick += ListResults_MouseClick;

            // Enable column click for sorting
            listResults.ColumnClick += ListResults_ColumnClick;

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
                        @"CREATE TABLE IF NOT EXISTS Games (
                                Id TEXT PRIMARY KEY,
                                Title TEXT NOT NULL,
                                Platform TEXT NOT NULL,
                                Link TEXT,
                                LastUpdated TEXT NOT NULL
                            )");
                }
                else
                {
                    // Ensure Platform column exists if using an older database
                    using var connection = GetConnection();
                    connection.Open();

                    // Check if Platform column exists
                    var tableInfo = connection.Query("PRAGMA table_info(Games)").ToList();
                    bool platformColumnExists = tableInfo.Any(row => (string)((IDictionary<string, object>)row)["name"] == "Platform");
                    bool linkColumnExists = tableInfo.Any(row => (string)((IDictionary<string, object>)row)["name"] == "Link");

                    // Add columns if missing
                    if (!platformColumnExists)
                    {
                        connection.Execute("ALTER TABLE Games ADD COLUMN Platform TEXT DEFAULT 'PS4'");
                    }

                    if (!linkColumnExists)
                    {
                        connection.Execute("ALTER TABLE Games ADD COLUMN Link TEXT");
                    }
                }
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
                List<CusaInfo> gameInfoList = new();

                // Extract CUSA and PPSA IDs from folder names and get dates
                foreach (string dir in directories)
                {
                    string dirName = Path.GetFileName(dir);
                    Match matchCusa = Regex.Match(dirName, @"CUSA\d{5}", RegexOptions.IgnoreCase);
                    Match matchPpsa = Regex.Match(dirName, @"PPSA\d{5}", RegexOptions.IgnoreCase);

                    if (matchCusa.Success)
                    {
                        gameInfoList.Add(new CusaInfo
                        {
                            CusaId = matchCusa.Value.ToUpper(),
                            Directory = dir,
                            Date = GetFolderDate(dir),
                            Platform = "PS4"
                        });
                    }
                    else if (matchPpsa.Success)
                    {
                        gameInfoList.Add(new CusaInfo
                        {
                            CusaId = matchPpsa.Value.ToUpper(),
                            Directory = dir,
                            Date = GetFolderDate(dir),
                            Platform = "PS5"
                        });
                    }
                }

                // Update status
                lblStatus.Text = $"Found {gameInfoList.Count} game IDs. Looking up titles...";

                // First check the database for cached titles
                using (var connection = GetConnection())
                {
                    connection.Open();
                    var cachedInfo = connection.Query<CusaTitle>(
                        "SELECT Id, Title, Platform, Link FROM Games WHERE Id IN @GameIds",
                        new { GameIds = gameInfoList.Select(c => c.CusaId).ToArray() }
                    ).ToDictionary(c => c.Id);

                    // Add cached info to the list
                    foreach (var info in gameInfoList.Where(c => cachedInfo.ContainsKey(c.CusaId)))
                    {
                        info.Title = cachedInfo[info.CusaId].Title;
                        info.Platform = cachedInfo[info.CusaId].Platform;
                        info.Link = cachedInfo[info.CusaId].Link;
                    }
                }

                // Lookup titles for items not in the database
                int processed = 0;
                int totalToProcess = gameInfoList.Count(c => string.IsNullOrEmpty(c.Title));
                int totalItems = gameInfoList.Count;

                foreach (CusaInfo info in gameInfoList.Where(c => string.IsNullOrEmpty(c.Title)))
                {
                    processed++;
                    lblStatus.Text = $"Processing: {processed}/{totalToProcess} (Total: {totalItems})";
                    Application.DoEvents();

                    await LookupAndSaveTitle(info);
                }

                // Display all results
                foreach (CusaInfo info in gameInfoList.OrderBy(c => c.CusaId))
                {
                    ListViewItem item = new(info.CusaId);

                    // Add title
                    string titleText = info.Title ?? "Not found";
                    item.SubItems.Add(titleText);

                    // Add platform
                    item.SubItems.Add(info.Platform);

                    // Add date
                    item.SubItems.Add(info.Date?.ToString("yyyy-MM-dd HH:mm:ss") ?? "Unknown");

                    // Add actions column with Open, Visit and Delete buttons
                    string actionsText = "📁 Open | 🔗 Visit | ❌ Delete";
                    item.SubItems.Add(actionsText);

                    // Store the info in the item's tag
                    item.Tag = info;

                    listResults.Items.Add(item);
                }

                lblStatus.Text = $"Completed. Found {gameInfoList.Count} game IDs.";
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
            // Determine platform based on ID prefix
            info.Platform = info.CusaId.StartsWith("CUSA", StringComparison.OrdinalIgnoreCase) ? "PS4" :
                            info.CusaId.StartsWith("PPSA", StringComparison.OrdinalIgnoreCase) ? "PS5" :
                            "Unknown";

            // Lookup title from appropriate source based on platform
            string? title = null;
            string? link = null;

            if (info.Platform == "PS4")
            {
                // First try orbispatches.com for PS4 games
                var orbisResult = await GetInfoFromOrbisPatches(info.CusaId);
                if (!string.IsNullOrEmpty(orbisResult.Title))
                {
                    title = orbisResult.Title;
                    link = orbisResult.Link;
                }

                // If not found, try serialstation.com
                if (string.IsNullOrEmpty(title))
                {
                    var serialResult = await GetInfoFromSerialStation(info.CusaId);
                    if (!string.IsNullOrEmpty(serialResult.Title))
                    {
                        title = serialResult.Title;
                        link = serialResult.Link;
                    }
                }
            }
            else if (info.Platform == "PS5")
            {
                // Use prosperopatches.com for PS5 games
                var prosperoResult = await GetInfoFromProsperoPatches(info.CusaId);
                if (!string.IsNullOrEmpty(prosperoResult.Title))
                {
                    title = prosperoResult.Title;
                    link = prosperoResult.Link;
                }
            }

            info.Title = title ?? string.Empty;
            info.Link = link ?? string.Empty;

            // Save to database if title was found
            if (!string.IsNullOrEmpty(title))
            {
                try
                {
                    using var connection = GetConnection();
                    connection.Open();
                    connection.Execute(
                        @"INSERT OR REPLACE INTO Games (Id, Title, Platform, Link, LastUpdated) 
                              VALUES (@Id, @Title, @Platform, @Link, @LastUpdated)",
                        new
                        {
                            Id = info.CusaId,
                            Title = info.Title,
                            Platform = info.Platform,
                            Link = info.Link,
                            LastUpdated = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                        });
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving to database: {ex.Message}");
                }
            }
        }

        private async Task<(string Title, string Link)> GetInfoFromOrbisPatches(string cusaId)
        {
            try
            {
                string url = $"https://orbispatches.com/{cusaId}";
                HtmlWeb web = new()
                {
                    AutoDetectEncoding = true,
                    OverrideEncoding = Encoding.UTF8
                };
                HtmlAgilityPack.HtmlDocument doc = await Task.Run(() => web.Load(url));

                string? title = null;

                // Parse title
                HtmlNode titleNode = doc.DocumentNode.SelectSingleNode("//header//h1[@class='bd-title']");
                if (titleNode != null)
                {
                    title = HttpUtility.HtmlDecode(titleNode.InnerText.Trim());
                    title = CleanupTitle(title);

                    if (!string.IsNullOrEmpty(title))
                        return (title, url);
                }

                return (string.Empty, string.Empty);
            }
            catch
            {
                return (string.Empty, string.Empty);
            }
        }

        private async Task<(string Title, string Link)> GetInfoFromSerialStation(string cusaId)
        {
            try
            {
                // Split CUSA ID for serialstation format
                string mainPart = cusaId[..4];
                string numberPart = cusaId[4..];

                string url = $"https://serialstation.com/titles/{mainPart}/{numberPart}";
                HtmlWeb web = new()
                {
                    AutoDetectEncoding = true,
                    OverrideEncoding = Encoding.UTF8
                };
                HtmlAgilityPack.HtmlDocument doc = await Task.Run(() => web.Load(url));

                string? title = null;

                HtmlNode titleNode = doc.DocumentNode.SelectSingleNode("//main[contains(@class,'container')]//h1");
                if (titleNode != null)
                {
                    title = HttpUtility.HtmlDecode(titleNode.InnerText.Trim());
                    title = CleanupTitle(title);

                    if (!string.IsNullOrEmpty(title))
                        return (title, url);
                }

                return (string.Empty, string.Empty);
            }
            catch
            {
                return (string.Empty, string.Empty);
            }
        }

        private static async Task<(string Title, string Link)> GetInfoFromProsperoPatches(string ppsaId)
        {
            try
            {
                string url = $"https://prosperopatches.com/{ppsaId}";
                HtmlWeb web = new()
                {
                    AutoDetectEncoding = true,
                    OverrideEncoding = Encoding.UTF8
                };
                HtmlAgilityPack.HtmlDocument doc = await Task.Run(() => web.Load(url));

                string? title = null;

                // Try multiple selectors to find the title
                HtmlNode? titleNode = doc.DocumentNode.SelectSingleNode("//header//h1[@class='bd-title']");

                if (titleNode == null)
                    titleNode = doc.DocumentNode.SelectSingleNode("//div[@class='container']//h1");

                if (titleNode == null)
                    titleNode = doc.DocumentNode.SelectSingleNode("//h1");

                if (titleNode == null)
                    titleNode = doc.DocumentNode.SelectSingleNode("//title");

                if (titleNode != null)
                {
                    title = HttpUtility.HtmlDecode(titleNode.InnerText.Trim());

                    if (title.Contains(ppsaId))
                    {
                        title = title.Replace(ppsaId, "").Trim();
                        title = title.Replace("- prosperopatches.com", "").Trim();
                    }

                    title = CleanupTitle(title);

                    if (!string.IsNullOrEmpty(title))
                        return (title, url);
                }

                return (string.Empty, string.Empty);
            }
            catch
            {
                return (string.Empty, string.Empty);
            }
        }

        private static string CleanupTitle(string title)
        {
            if (string.IsNullOrEmpty(title))
                return title;

            // Remove any trailing unnecessary text that might appear in titles
            title = Regex.Replace(title, @"\s*\(CUSA\d+\)\s*$", "", RegexOptions.IgnoreCase);

            // Remove any duplicate spaces
            title = Regex.Replace(title, @"\s+", " ");

            return title.Trim();
        }

        private static DateTime? GetFolderDate(string folderPath)
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
                // Handle Actions column (column index 4)
                if (columnIndex == 4)
                {
                    string actionText = hitTest.SubItem.Text;

                    // Calculate positions of components in the Actions text
                    int firstSeparatorIndex = actionText.IndexOf("|");
                    int secondSeparatorIndex = actionText.LastIndexOf("|");

                    // Calculate widths for hit testing
                    int openWidth = TextRenderer.MeasureText(actionText.Substring(0, firstSeparatorIndex), listResults.Font).Width;

                    int visitWidth = 0;
                    if (firstSeparatorIndex != secondSeparatorIndex) // If we have a Visit button
                    {
                        visitWidth = TextRenderer.MeasureText(
                            actionText.Substring(0, secondSeparatorIndex),
                            listResults.Font).Width;
                    }
                    else
                    {
                        visitWidth = openWidth; // If no Visit button, delete starts right after open
                    }

                    int clickPosition = e.X - hitTest.SubItem.Bounds.Left;

                    // Open folder action
                    if (clickPosition <= openWidth)
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
                    // Visit link action - only if we have a link
                    else if (firstSeparatorIndex != secondSeparatorIndex && clickPosition <= visitWidth && !string.IsNullOrEmpty(info.Link))
                    {
                        try
                        {
                            // Open the link in default browser
                            Process.Start(new ProcessStartInfo
                            {
                                FileName = info.Link,
                                UseShellExecute = true
                            });
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Error opening link: {ex.Message}", "Error",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    // Delete action
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
                    using var connection = GetConnection();
                    connection.Open();
                    int affected = connection.Execute(
                        "DELETE FROM Games WHERE Id = @Id",
                        new { Id = cusaId });

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
                catch (Exception ex)
                {
                    MessageBox.Show($"Error deleting item: {ex.Message}", "Database Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // Column sorting handler for ListView
        private void ListResults_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Don't sort the Actions column (column index 4)
            if (e.Column == 4)
                return;

            ListView listView = (ListView)sender;

            // Create or get the sorter
            ListViewColumnSorter sorter;
            if (listView.ListViewItemSorter == null)
            {
                sorter = new ListViewColumnSorter();
                listView.ListViewItemSorter = sorter;
            }
            else
            {
                sorter = (ListViewColumnSorter)listView.ListViewItemSorter;
            }

            // Set the column and update sort direction
            if (sorter.SortColumn == e.Column)
            {
                // Reverse the sort direction if clicking the same column
                sorter.SortOrder = sorter.SortOrder == SortOrder.Ascending
                    ? SortOrder.Descending
                    : SortOrder.Ascending;
            }
            else
            {
                // Set the new sort column and default to ascending
                sorter.SortColumn = e.Column;
                sorter.SortOrder = SortOrder.Ascending;
            }

            // Perform the sort
            listView.Sort();
        }
    }

    public class CusaInfo
    {
        public string CusaId { get; set; }
        public string Title { get; set; }
        public string Directory { get; set; }
        public DateTime? Date { get; set; }
        public string Platform { get; set; }
        public string Link { get; set; }
    }

    public class CusaTitle
    {
        public string Id { get; set; }
        public string Title { get; set; }
        public string Platform { get; set; }
        public string Link { get; set; }
        public string LastUpdated { get; set; }
    }

    // Implements the column sorting for ListView
    public class ListViewColumnSorter : IComparer
    {
        // Column to sort
        public int SortColumn { get; set; }

        // Sort order
        public SortOrder SortOrder { get; set; }

        // Case insensitive comparer
        private readonly CaseInsensitiveComparer _objectCompare;

        // Constructor
        public ListViewColumnSorter()
        {
            // Initialize with default values
            SortColumn = 0;
            SortOrder = SortOrder.Ascending;
            _objectCompare = new CaseInsensitiveComparer();
        }

        // Comparison method implementation
        public int Compare(object x, object y)
        {
            // Convert the objects to list view items
            ListViewItem listViewX = (ListViewItem)x;
            ListViewItem listViewY = (ListViewItem)y;

            // Get text values to compare
            string textX = listViewX.SubItems[SortColumn].Text;
            string textY = listViewY.SubItems[SortColumn].Text;

            // Date column needs special handling (column index 3)
            if (SortColumn == 3)
            {
                // Try to parse as dates
                if (DateTime.TryParse(textX, out DateTime dateX) && DateTime.TryParse(textY, out DateTime dateY))
                {
                    return SortOrder == SortOrder.Ascending ?
                        DateTime.Compare(dateX, dateY) :
                        DateTime.Compare(dateY, dateX);
                }
            }

            // Perform the comparison for other columns
            int compareResult = _objectCompare.Compare(textX, textY);

            // Return the result based on sort order
            return SortOrder == SortOrder.Ascending ? compareResult : -compareResult;
        }
    }
}
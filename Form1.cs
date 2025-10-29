using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ExcelViewer
{
    public partial class Form1 : Form
    {
        // --- State ---
        private DataTable? _data;                 // In-memory client data
        private string? _currentFilePath;         // Selected file path

        // --- UI: top menu ---
        private MenuStrip menu = default!;
        private ToolStripMenuItem mFile = default!;
        private ToolStripMenuItem mClientData = default!;
        private ToolStripMenuItem mMapping = default!;
        private ToolStripMenuItem mValidation = default!;
        private ToolStripMenuItem mUdData = default!;
        private ToolStripMenuItem mAbout = default!;

        // --- UI: tabs ---
        private TabControl tabs = default!;
        private TabPage tabClientData = default!;
        private TabPage tabMapping = default!;
        private TabPage tabValidation = default!;
        private TabPage tabUdData = default!;
        private TabPage tabAbout = default!;

        // --- Client Data tab controls ---
        private Button btnImport = default!;
        private Label lblFileCaption = default!;
        private Label lblFileName = default!;
        private DataGridView gridPreview = default!;

        // --- Mapping tab controls (skeleton) ---
        private DataGridView gridMapping = default!;
        private Button btnSaveMapping = default!;

        // --- Validation tab controls ---
        private Label lblValidationHeader = default!;
        private Label lblDupStatus = default!;
        private Label lblMissingStatus = default!;
        private DataGridView gridDuplicates = default!;
        private DataGridView gridMissing = default!;

        // --- UD Data tab controls (placeholder) ---
        private Label lblUdHeader = default!;

        // --- About tab ---
        private Label lblAbout = default!;

        public Form1()
        {
            InitializeComponent();
            BuildUi();
        }

        private void BuildUi()
        {
            Text = "Client Data Import Wizard";
            Width = 1200;
            Height = 800;
            StartPosition = FormStartPosition.CenterScreen;

            // Menu
            menu = new MenuStrip();

            mFile = new ToolStripMenuItem("File");
            var miNewImport = new ToolStripMenuItem("New Import", null, (_, __) => NewImport());
            var miOpen = new ToolStripMenuItem("Open...", null, (_, __) => ImportExcel());
            var miSave = new ToolStripMenuItem("Save", null, (_, __) => MessageBox.Show("Save not implemented yet."));
            var miSaveAs = new ToolStripMenuItem("Save As...", null, (_, __) => MessageBox.Show("Save As not implemented yet."));
            mFile.DropDownItems.AddRange(new ToolStripItem[] { miNewImport, miOpen, new ToolStripSeparator(), miSave, miSaveAs });

            mClientData = new ToolStripMenuItem("Client Data", null, (_, __) => tabs.SelectedTab = tabClientData);
            mMapping    = new ToolStripMenuItem("Mapping",     null, (_, __) => tabs.SelectedTab = tabMapping);
            mValidation = new ToolStripMenuItem("Validation",  null, (_, __) => tabs.SelectedTab = tabValidation);
            mUdData     = new ToolStripMenuItem("UD Data",     null, (_, __) => tabs.SelectedTab = tabUdData);
            mAbout      = new ToolStripMenuItem("About",       null, (_, __) => tabs.SelectedTab = tabAbout);

            menu.Items.AddRange(new ToolStripItem[] { mFile, mClientData, mMapping, mValidation, mUdData, mAbout });
            MainMenuStrip = menu;
            Controls.Add(menu);

            // Tabs
            tabs = new TabControl { Dock = DockStyle.Fill };
            tabClientData = new TabPage("Client Data");
            tabMapping    = new TabPage("Mapping");
            tabValidation = new TabPage("Validation");
            tabUdData     = new TabPage("UD Data");
            tabAbout      = new TabPage("About");
            tabs.TabPages.AddRange(new[] { tabClientData, tabMapping, tabValidation, tabUdData, tabAbout });
            Controls.Add(tabs);

            BuildClientDataTab();
            BuildMappingTab();
            BuildValidationTab();
            BuildUdTab();
            BuildAboutTab();
        }

        // ---------------- Client Data ----------------
        private void BuildClientDataTab()
        {
            // Top strip: Import + File name
            var top = new FlowLayoutPanel
            {
                Dock = DockStyle.Top,
                Height = 56,
                Padding = new Padding(10),
                FlowDirection = FlowDirection.LeftToRight
            };

            btnImport = new Button { Text = "Import", AutoSize = true };
            btnImport.Click += (_, __) => ImportExcel();

            lblFileCaption = new Label { Text = "File name:", AutoSize = true, Margin = new Padding(16, 8, 4, 4) };
            lblFileName = new Label { Text = "—", AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold), Margin = new Padding(4, 8, 4, 4) };

            top.Controls.Add(btnImport);
            top.Controls.Add(lblFileCaption);
            top.Controls.Add(lblFileName);

            // Preview grid
            gridPreview = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells,
                AutoGenerateColumns = true
            };

            tabClientData.Controls.Add(gridPreview);
            tabClientData.Controls.Add(top);
        }

        // ---------------- Mapping (skeleton) ----------------
        private void BuildMappingTab()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 2
            };
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 50));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var header = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(10),
                FlowDirection = FlowDirection.LeftToRight
            };
            var lbl = new Label { Text = "Fields Mapping", AutoSize = true, Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold) };
            btnSaveMapping = new Button { Text = "Save Mapping", AutoSize = true, Margin = new Padding(16, 4, 4, 4) };
            btnSaveMapping.Click += (_, __) => MessageBox.Show("Mapping persistence not implemented yet.");
            header.Controls.Add(lbl);
            header.Controls.Add(btnSaveMapping);

            gridMapping = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = false,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoGenerateColumns = false
            };

            // Two-column skeleton: Source Column (from Excel) + Target Field (your system)
            var colSource = new DataGridViewTextBoxColumn { HeaderText = "Source Column", DataPropertyName = "Source", Width = 250, ReadOnly = true };
            var colTarget = new DataGridViewComboBoxColumn { HeaderText = "Target Field", DataPropertyName = "Target", Width = 250 };
            // Example targets; replace with real system fields
            ((DataGridViewComboBoxColumn)colTarget).Items.AddRange(new[] { "ID", "FirstName", "LastName", "Email", "Phone", "Custom1", "Custom2" });

            gridMapping.Columns.Add(colSource);
            gridMapping.Columns.Add(colTarget);

            panel.Controls.Add(header, 0, 0);
            panel.Controls.Add(gridMapping, 0, 1);
            tabMapping.Controls.Add(panel);
        }

        // ---------------- Validation ----------------
        private void BuildValidationTab()
        {
            var panel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 3,
                Padding = new Padding(10)
            };
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50));
            panel.RowStyles.Add(new RowStyle(SizeType.Absolute, 36));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 50));
            panel.RowStyles.Add(new RowStyle(SizeType.Percent, 50));

            lblValidationHeader = new Label
            {
                Text = "Data Validation",
                AutoSize = true,
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold),
                Dock = DockStyle.Fill
            };
            panel.SetColumnSpan(lblValidationHeader, 2);

            lblDupStatus = new Label { Text = "Duplicate Records: —", AutoSize = true, Dock = DockStyle.Fill, Padding = new Padding(0, 8, 0, 8) };
            lblMissingStatus = new Label { Text = "Missing IDs: —", AutoSize = true, Dock = DockStyle.Fill, Padding = new Padding(0, 8, 0, 8) };

            gridDuplicates = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
            };

            gridMissing = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells
            };

            panel.Controls.Add(lblValidationHeader, 0, 0);
            panel.Controls.Add(lblDupStatus, 0, 1);
            panel.Controls.Add(lblMissingStatus, 1, 1);
            panel.Controls.Add(gridDuplicates, 0, 2);
            panel.Controls.Add(gridMissing, 1, 2);

            tabValidation.Controls.Add(panel);
        }

        // ---------------- UD Data (placeholder) ----------------
        private void BuildUdTab()
        {
            lblUdHeader = new Label
            {
                Text = "UD Data (for <client file name>)",
                AutoSize = true,
                Dock = DockStyle.Top,
                Padding = new Padding(10),
                Font = new System.Drawing.Font("Segoe UI", 10, System.Drawing.FontStyle.Bold)
            };
            tabUdData.Controls.Add(lblUdHeader);
        }

        // ---------------- About ----------------
        private void BuildAboutTab()
        {
            lblAbout = new Label
            {
                Text = "About\n\nClient Data Import Wizard\nBuilt with .NET (WinForms) + ClosedXML.",
                AutoSize = true,
                Dock = DockStyle.Top,
                Padding = new Padding(10)
            };
            tabAbout.Controls.Add(lblAbout);
        }

        // ---------------- Actions ----------------

        private void NewImport()
        {
            _data = null;
            _currentFilePath = null;
            lblFileName.Text = "—";
            gridPreview.DataSource = null;
            gridMapping.DataSource = null;
            gridDuplicates.DataSource = null;
            gridMissing.DataSource = null;
            lblDupStatus.Text = "Duplicate Records: —";
            lblMissingStatus.Text = "Missing IDs: —";
            lblUdHeader.Text = "UD Data (for <client file name>)";
            tabs.SelectedTab = tabClientData;
        }

        private void ImportExcel()
        {
            using var ofd = new OpenFileDialog
            {
                Title = "Select client Excel file",
                Filter = "Excel Files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All Files (*.*)|*.*",
                Multiselect = false
            };

            if (ofd.ShowDialog(this) != DialogResult.OK) return;

            try
            {
                _currentFilePath = ofd.FileName;
                lblFileName.Text = Path.GetFileName(_currentFilePath);
                lblUdHeader.Text = $"UD Data for {lblFileName.Text}";

                _data = ReadExcelToDataTable(_currentFilePath);
                gridPreview.DataSource = _data;

                // Seed mapping rows from current columns
                var mapRows = new List<MappingRow>();
                foreach (DataColumn col in _data.Columns)
                {
                    mapRows.Add(new MappingRow { Source = col.ColumnName, Target = "" });
                }
                gridMapping.DataSource = mapRows;

                // Run validation right away
                RunValidation();
                tabs.SelectedTab = tabMapping; // move forward
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Failed to load Excel:\n\n" + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RunValidation()
        {
            if (_data == null || _data.Columns.Count == 0)
            {
                lblDupStatus.Text = "Duplicate Records: —";
                lblMissingStatus.Text = "Missing IDs: —";
                gridDuplicates.DataSource = null;
                gridMissing.DataSource = null;
                return;
            }

            // Choose ID column: prefer a column literally named "ID" (case-insensitive); otherwise first column
            var idCol = _data.Columns.Cast<DataColumn>()
                          .FirstOrDefault(c => string.Equals(c.ColumnName, "ID", StringComparison.OrdinalIgnoreCase))
                        ?? _data.Columns[0];

            // Missing IDs: null/empty/whitespace
            var missing = _data.AsEnumerable()
                               .Where(r => string.IsNullOrWhiteSpace(r[idCol].ToString()))
                               .CopyToDataTableOrNull();

            if (missing == null)
            {
                lblMissingStatus.Text = "Missing IDs: ✓ None found";
                gridMissing.DataSource = null;
            }
            else
            {
                lblMissingStatus.Text = $"Missing IDs: {missing.Rows.Count} row(s)";
                gridMissing.DataSource = missing;
            }

            // Duplicates: group by ID where not empty, count > 1
            var notEmpty = _data.AsEnumerable()
                                .Where(r => !string.IsNullOrWhiteSpace(r[idCol].ToString()));

            var dupKeys = notEmpty.GroupBy(r => r[idCol].ToString()!)
                                  .Where(g => g.Count() > 1)
                                  .Select(g => g.Key)
                                  .ToHashSet();

            if (dupKeys.Count == 0)
            {
                lblDupStatus.Text = "Duplicate Records: ✓ None found";
                gridDuplicates.DataSource = null;
            }
            else
            {
                var dupRows = notEmpty.Where(r => dupKeys.Contains(r[idCol].ToString()!))
                                      .CopyToDataTableOrNull();
                lblDupStatus.Text = $"Duplicate Records: {dupRows?.Rows.Count ?? 0} row(s) across {dupKeys.Count} duplicated ID(s)";
                gridDuplicates.DataSource = dupRows;
            }
        }

        // ---------------- Helpers ----------------

        private static DataTable ReadExcelToDataTable(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException("File not found", path);

            using var wb = new XLWorkbook(path);
            var ws = wb.Worksheets.FirstOrDefault()
                     ?? throw new InvalidOperationException("No worksheets found in the Excel file.");

            var used = ws.RangeUsed()
                       ?? throw new InvalidOperationException("The worksheet is empty.");

            var dt = new DataTable();

            int headerRow = used.FirstRowUsed().RowNumber();
            int firstCol  = used.FirstColumnUsed().ColumnNumber();
            int lastCol   = used.LastColumnUsed().ColumnNumber();
            int lastRow   = used.LastRowUsed().RowNumber();

            // Headers
            for (int col = firstCol; col <= lastCol; col++)
            {
                var header = ws.Cell(headerRow, col).GetString();
                if (string.IsNullOrWhiteSpace(header))
                    header = $"Column{col - firstCol + 1}";

                var unique = header;
                int dup = 1;
                while (dt.Columns.Contains(unique))
                    unique = header + "_" + (++dup);

                dt.Columns.Add(unique);
            }

            // Rows
            for (int row = headerRow + 1; row <= lastRow; row++)
            {
                var newRow = dt.NewRow();
                int colIndex = 0;
                for (int col = firstCol; col <= lastCol; col++)
                {
                    newRow[colIndex++] = ws.Cell(row, col).GetString();
                }
                dt.Rows.Add(newRow);
            }

            return dt;
        }
    }

    // Simple mapping row model for the Mapping grid
    internal class MappingRow
    {
        public string Source { get; set; } = "";
        public string Target { get; set; } = "";
    }

    // Small helper to avoid exceptions when sequence is empty
    internal static class DataTableExtensions
    {
        public static DataTable? CopyToDataTableOrNull(this IEnumerable<DataRow> rows)
        {
            using var e = rows.GetEnumerator();
            if (!e.MoveNext()) return null; // empty
            // Build a table with the same schema as the first row's table
            var table = e.Current!.Table.Clone();
            table.ImportRow(e.Current);
            while (e.MoveNext())
                table.ImportRow(e.Current!);
            return table;
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Library.Forms;
using System.Data.Entity;
using System.Data.Odbc;
using System.Globalization;
using System.Diagnostics;

namespace SO_Delivery_Interface
{
    public partial class MainForm : Form
    {
        List<string> customers = new List<string>();
        List<string> parts = new List<string>();
        private static string DSN = "jobboss32";
        private static string userName = "jbread";
        private static string password = "Cloudy2Day";

        public MainForm()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // set up date time picker1
            idateDateTimePicker.Value = DateTime.Today.AddDays(-1);
            idateDateTimePicker.MinDate = new DateTime(2015, 7, 29);
            idateDateTimePicker.CloseUp += this.UpdateDateLabels;
            idateDateTimePicker.CloseUp += this.Update_Monthly;
            idateDateTimePicker.CloseUp += this.Update_30d90d6m1yTable;
            idateDateTimePicker.CloseUp += this.UpdateMainTable;

            // set up date time picker2
            comparisondateDateTimePicker.Value = DateTime.Today;
            comparisondateDateTimePicker.MinDate = new DateTime(2015, 7, 29);
            comparisondateDateTimePicker.CloseUp += this.UpdateDateLabels;
            comparisondateDateTimePicker.CloseUp += this.Update_Monthly;
            comparisondateDateTimePicker.CloseUp += this.Update_30d90d6m1yTable;
            comparisondateDateTimePicker.CloseUp += this.UpdateMainTable;

            // default radio button for datagridview
            openLinesRadioButton.Checked = true;
            openLinesRadioButton.CheckedChanged += this.UpdateMainTable;

            // default radio button for due dates
            partsRadioButton.Checked = true;
            partsRadioButton.CheckedChanged += this.Update_30d90d6m1yTable;
            partsRadioButton.CheckedChanged += this.Update_Monthly;

            // Set up datagrid view
            UpdateMainTable(new Object(), new EventArgs());
            linesDataGridView.CellClick += this.Update_Monthly;
            linesDataGridView.CellClick += this.Update_30d90d6m1yTable;
            linesDataGridView.CellClick += this.updateLabels;
            linesDataGridView.Columns[7].DefaultCellStyle.Format = "c";
            linesDataGridView.Columns[8].DefaultCellStyle.Format = "c";

            // load up drowpdown
            filterDropDown.Items.Add("Customer");
            filterDropDown.Items.Add("Part No.");
            filterDropDown.SelectedIndex = 0;

            // set up dataGridView_idate
            dataGridView_iDate.Rows.Add(5);
            dataGridView_iDate.Rows[0].HeaderCell.Value = "pd";
            dataGridView_iDate.Rows[1].HeaderCell.Value = "30d";
            dataGridView_iDate.Rows[2].HeaderCell.Value = "90d";
            dataGridView_iDate.Rows[3].HeaderCell.Value = "6mo";
            dataGridView_iDate.Rows[4].HeaderCell.Value = "1yr";
            dataGridView_iDate.RowHeadersDefaultCellStyle.Padding = new Padding(linesDataGridView.RowHeadersWidth);
            dataGridView_iDate.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dataGridView_iDate_RowPostPaint);
            dataGridView_iDate.ClearSelection();
            dataGridView_iDate.SelectionChanged += this.dgvSomeDataGridView_SelectionChanged;

            // set up dataGridView_current
            dataGridView_current.Rows.Add(5);
            dataGridView_current.Rows[0].HeaderCell.Value = "pd";
            dataGridView_current.Rows[1].HeaderCell.Value = "30d";
            dataGridView_current.Rows[2].HeaderCell.Value = "90d";
            dataGridView_current.Rows[3].HeaderCell.Value = "6mo";
            dataGridView_current.Rows[4].HeaderCell.Value = "1yr";
            dataGridView_current.RowHeadersDefaultCellStyle.Padding = new Padding(linesDataGridView.RowHeadersWidth);
            dataGridView_current.RowPostPaint += new DataGridViewRowPostPaintEventHandler(dataGridView_current_RowPostPaint);
            dataGridView_current.ClearSelection();
            dataGridView_current.SelectionChanged += this.dgvSomeDataGridView_SelectionChanged;

            // set up iDate_year dropdown
            for (int i = 2011; i <= DateTime.Now.Year + 3; i++)
                iDate_yearDropDown.Items.Add(i.ToString());
            for (int i = 0; i < iDate_yearDropDown.Items.Count; i++)
            {
                if (iDate_yearDropDown.Items[i].ToString().Equals(DateTime.Now.Year.ToString()))
                {
                    iDate_yearDropDown.SelectedIndex = i;
                    break;
                }
            }
            iDate_yearDropDown.SelectedIndexChanged += this.Update_Monthly;

            // set up quarters_iDate
            monthly_iDateGridView.Rows.Add(12);
            for (int i = 1; i <= 12; i++)
                monthly_iDateGridView.Rows[i - 1].HeaderCell.Value = UpperCaseFirst(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i).Substring(0, 3));
            monthly_iDateGridView.RowHeadersDefaultCellStyle.Padding = new Padding(linesDataGridView.RowHeadersWidth);
            monthly_iDateGridView.RowPostPaint += new DataGridViewRowPostPaintEventHandler(quarter_iDateGridView_RowPostPaint);

            // set up quarters_current
            monthly_currentGridView.Rows.Add(12);
            for (int i = 1; i <= 12; i++)
                monthly_currentGridView.Rows[i - 1].HeaderCell.Value = UpperCaseFirst(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i).Substring(0, 3));
            monthly_currentGridView.RowHeadersDefaultCellStyle.Padding = new Padding(linesDataGridView.RowHeadersWidth);
            monthly_currentGridView.RowPostPaint += new DataGridViewRowPostPaintEventHandler(quarter_currentGridView_RowPostPaint);
            monthly_currentGridView.Rows[DateTime.Now.Month - 1].Selected = true;

            // set up date
            currentDateLabel.Text = comparisondateDateTimePicker.Value.ToShortDateString();
            date2Label.Text = comparisondateDateTimePicker.Value.ToShortDateString();

            // set up iDate
            iDateLabel.Text = idateDateTimePicker.Value.ToShortDateString();
            date1Label.Text = idateDateTimePicker.Value.ToShortDateString();

            // set up firm checkbox
            firmCostCheckBox.Checked = true;
            firmCostCheckBox.CheckedChanged += this.Update_Monthly;
            firmCostCheckBox.CheckedChanged += this.Update_30d90d6m1yTable;

            // set up forecast checkbox
            forecastCostCheckBox.Checked = true;
            forecastCostCheckBox.CheckedChanged += this.Update_Monthly;
            forecastCostCheckBox.CheckedChanged += this.Update_30d90d6m1yTable;

            // set up total checkbox
            shippedCostCheckBox.CheckedChanged += this.Update_Monthly;
            shippedCostCheckBox.CheckedChanged += this.Update_30d90d6m1yTable;

            // set up selectAll checkbox
            selectAllCostCheckBox.CheckedChanged += this.Update_Monthly;
            selectAllCostCheckBox.CheckedChanged += this.Update_30d90d6m1yTable;

            // set up addutton
            addButton.Click += this.UpdateMainTable;

            // set up remove button
            removeButton.Click += this.UpdateMainTable;
        }

        private static string UpperCaseFirst(string s)
        {
            // Check for empty string.
            if (string.IsNullOrEmpty(s))
            {
                return string.Empty;
            }
            // Return char and concat substring.
            return char.ToUpper(s[0]) + s.Substring(1);
        }

        private void UpdateMainTable(object sender, EventArgs e)
        {

            if (openLinesRadioButton.Checked)
            {
                using (OdbcConnection con = new OdbcConnection("DSN=" + DSN + ";UID=" + userName + ";PWD=" + password))
                {
                    string query = "SELECT Sales_Order AS 'Sales Order', SO_Line AS 'SO Line', Customer, Material, Customer_PO AS 'Customer PO', Order_Qty AS 'Order Qty', Promised_Date AS 'Promised Date', Unit_Price AS 'Unit Price', Total_Price AS 'Total Price', Status, Last_Updated AS 'Last Updated', idate, Rev, Description\n" +
                                    "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                                    "WHERE Status NOT IN ('Removed', 'Shipped', 'Closed') AND ( " + (customers.Count == 0 ? "1 = 1" : "1= 0") + " OR Customer IN ('" + string.Join("','", customers.ToArray()) + "')) AND ( " + (parts.Count == 0 ? "1 = 1" : "1= 0") + " OR Material IN ('" + string.Join("','", parts.ToArray()) + "')) AND CONVERT(DATETIME, idate) =\n" +
                                        "\t(SELECT CONVERT(DATETIME,MAX(InnerT.idate))\n" +
                                        "\tFROM ATIDelivery.dbo.Delivery_Lines AS InnerT\n" +
                                        "\tWHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + idateDateTimePicker.Value.ToShortDateString() + "'))\n" +
                                    "ORDER BY Promised_Date, SO_Line;";
                    con.Open();

                    OdbcCommand com = new OdbcCommand(query, con);
                    try
                    {
                        OdbcDataReader reader = com.ExecuteReader();

                        DataSet ds = new DataSet();
                        DataTable dt = new DataTable();
                        ds.Tables.Add(dt);
                        ds.Load(reader, LoadOption.PreserveChanges, ds.Tables[0]);

                        linesDataGridView.DataSource = ds.Tables[0];
                        linesDataGridView.Columns[12].Visible = false;
                        linesDataGridView.Columns[13].Visible = false;

                        // resize columns
                        linesDataGridView.Columns[0].Width = 85;
                        linesDataGridView.Columns[1].Width = 70;
                        linesDataGridView.Columns[2].Width = 90;
                        linesDataGridView.Columns[3].Width = 130;
                        linesDataGridView.Columns[4].Width = 110;
                        linesDataGridView.Columns[5].Width = 70;
                        linesDataGridView.Columns[9].Width = 90;
                        linesDataGridView.Columns[10].Width = 120;
                        linesDataGridView.Columns[11].Width = 120;
                    }
                    catch (Exception er)
                    {
                        MessageBox.Show(er.Message);
                    }
                }
            }
            else // Show ALL
            {
                using (OdbcConnection con = new OdbcConnection("DSN=" + DSN + ";UID=" + userName + ";PWD=" + password))
                {
                    // this query returns all of the SO lines
                    string query = "SELECT Sales_Order AS 'Sales Order', SO_Line AS 'SO Line', Customer, Material, Customer_PO AS 'Customer PO', Order_Qty AS 'Order Qty', Promised_Date AS 'Promised Date', Unit_Price AS 'Unit Price', Total_Price AS 'Total Price', Status, Last_Updated AS 'Last Updated', idate, Rev, Description, pk, SO_Detail\n" +
                                    "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                                    "WHERE  idate >= '" + idateDateTimePicker.Value.ToShortDateString() + "' AND idate < '" + idateDateTimePicker.Value.AddDays(1).ToShortDateString() + "' AND ( " + (customers.Count == 0 ? "1 = 1" : "1= 0") + " OR Customer IN ('" + string.Join("','", customers.ToArray()) + "')) AND ( " + (parts.Count == 0 ? "1 = 1" : "1= 0") + " OR Material IN ('" + string.Join("','", parts.ToArray()) + "'))\n" +
                                    "ORDER BY Promised_Date, SO_Line;\n";
                    con.Open();

                    OdbcCommand com = new OdbcCommand(query, con);
                    OdbcDataReader reader = com.ExecuteReader();


                    DataSet ds = new DataSet();
                    DataTable dt = new DataTable();
                    ds.Tables.Add(dt);
                    ds.Load(reader, LoadOption.PreserveChanges, ds.Tables[0]);

                    linesDataGridView.DataSource = ds.Tables[0];

                    linesDataGridView.Columns[12].Visible = false;
                    linesDataGridView.Columns[13].Visible = false;
                    linesDataGridView.Columns[14].Visible = false;
                    linesDataGridView.Columns[15].Visible = false;

                    query = "SELECT	CASE WHEN readT.Sales_Order = compareT.Sales_Order THEN 1 ELSE 0 END AS 'Sales_Order Change',\n" +
                                "\tCASE WHEN readT.SO_Line = compareT.SO_Line THEN 1 ELSE 0 END AS 'SO_Line Change',\n" +
                                "\tCASE WHEN readT.Customer = compareT.Customer THEN 1 ELSE 0 END AS 'Customer Change',\n" +
                                "\tCASE WHEN readT.Material = compareT.Material THEN 1 ELSE 0 END AS 'Part # Change',\n" +
                                "\tCASE WHEN readT.Customer_PO = compareT.Customer_PO THEN 1 ELSE 0 END AS 'Customer_PO Change',\n" +
                                "\tCASE WHEN readT.Order_Qty = compareT.Order_Qty THEN 1 ELSE 0 END AS 'Order Qty Change',\n" +
                                "\tCASE WHEN readT.Promised_Date = compareT.Promised_Date THEN 1 ELSE 0 END AS 'Promised_Date Change',\n" +
                                "\tCASE WHEN readT.Unit_Price = compareT.Unit_Price THEN 1 ELSE 0 END AS 'Unit_Price Changed`',\n" +
                                "\tCASE WHEN readT.Total_Price = compareT.Total_Price THEN 1 ELSE 0 END AS 'Total Price Change',\n" +
                                "\tCASE WHEN readT.Status = compareT.Status THEN 1 ELSE 0 END AS 'Status Change',\n" +
                                "\tCASE WHEN readT.Last_Updated = compareT.Last_Updated THEN 1 ELSE 1 END AS 'Last_Updated Change',\n" +
                                "\tCASE WHEN readT.idate = compareT.idate THEN 1 ELSE 1 END AS 'idate Change',\n" +
                                "\tCASE WHEN readT.Rev = compareT.Rev THEN 1 ELSE 0 END AS 'Rev Change',\n" +
                                "\tCASE WHEN readT.Description = compareT.Description THEN 1 ELSE 0 END AS 'Description Change',\n" +
                                "\tCASE WHEN readT.pk = compareT.pk THEN 1 ELSE 1 END AS 'pk changed',\n" +
                                "\tCASE when readT.SO_Detail = compareT.SO_Detail THEN 1 ELSE 0 END AS 'SO_Detail changed'\n" +
                          "FROM ATIDelivery.dbo.Delivery_Lines AS readT\n" +
                          "LEFT JOIN\n" +
                                "\t(SELECT * FROM ATIDelivery.dbo.Delivery_Lines AS compareT2 WHERE compareT2.idate = (SELECT MAX(innerT.idate) FROM ATIDelivery.dbo.Delivery_Lines AS innerT WHERE  innerT.SO_Detail = compareT2.SO_Detail AND innerT.idate < CONVERT(DATE, '" + idateDateTimePicker.Value.ToShortDateString() + "'))) AS compareT\n" +
                          "ON compareT.SO_Detail = readT.SO_Detail\n" +
                          "WHERE  readT.idate >= '" + idateDateTimePicker.Value.ToShortDateString() + "' AND readT.idate < '" + idateDateTimePicker.Value.AddDays(1).ToShortDateString() + "'  AND ( " + (customers.Count == 0 ? "1 = 1" : "1= 0") + " OR readT.Customer IN ('" + string.Join("','", customers.ToArray()) + "')) AND ( " + (parts.Count == 0 ? "1 = 1" : "1= 0") + " OR readT.Material IN ('" + string.Join("','", parts.ToArray()) + "'))\n" +
                          "ORDER BY readT.Promised_Date, readT.SO_Line\n";

                    com = new OdbcCommand(query, con);
                    reader = com.ExecuteReader();

                    int row = 0;
                    try
                    {
                        while (reader.Read())
                        {
                            if (reader.GetInt32(15) == 0)
                            {
                                linesDataGridView.Rows[row].Cells[0].Style = new DataGridViewCellStyle { ForeColor = Color.FromArgb(21, 196, 21) };
                                row++;
                                continue;

                            }
                            for (int i = 0; i < 12; i++)
                            {
                                if (reader.GetInt32(i) == 0)
                                    linesDataGridView.Rows[row].Cells[i].Style = new DataGridViewCellStyle { ForeColor = Color.Blue };
                            }
                            row++;
                        }
                    }
                    catch (ArgumentOutOfRangeException ex)
                    {
                        MessageBox.Show("Exception: " + ex.Message + "\nRow: " + row);
                    }

                    // resize columns
                    linesDataGridView.Columns[0].Width = 85;
                    linesDataGridView.Columns[1].Width = 70;
                    linesDataGridView.Columns[2].Width = 90;
                    linesDataGridView.Columns[3].Width = 130;
                    linesDataGridView.Columns[4].Width = 110;
                    linesDataGridView.Columns[5].Width = 70;
                    linesDataGridView.Columns[9].Width = 90;
                    linesDataGridView.Columns[10].Width = 120;
                    linesDataGridView.Columns[11].Width = 120;

                }
            }

            rowCounterLabel.Text = linesDataGridView.RowCount.ToString();
        }

        private void addButton_Click(object sender, EventArgs e)
        {
            // check empty
            if (filterInput.TextLength == 0 || filterDropDown.SelectedItem == null)
                return;

            // check if drop down is set to customer or parts
            if (filterDropDown.SelectedItem.ToString().Equals("Customer"))
            {
                if (!customers.Exists(x => x.Equals(filterInput.Text.ToUpper())))
                {
                    customers.Add(filterInput.Text.ToUpper());
                    customersListBox.Items.Add(filterInput.Text.ToUpper());
                }
            }
            else if (filterDropDown.SelectedItem.ToString().Equals("Part No."))
            {
                if (!parts.Exists(x => x.Equals(filterInput.Text.ToUpper())))
                {
                    parts.Add(filterInput.Text.ToUpper());
                    partsListBox.Items.Add(filterInput.Text.ToUpper());
                }
            }
            filterInput.Clear();
        }

        private void removeButton_Click(object sender, EventArgs e)
        {
            if (customersListBox.SelectedItem == null && partsListBox.SelectedItem == null)
                return;

            int index = 0;

            while (index < customersListBox.Items.Count)
            {
                if (customersListBox.GetSelected(index))
                {
                    customers.Remove(customersListBox.Items[index].ToString());
                    customersListBox.Items.RemoveAt(index);
                }
                else
                    index++;
            }

            index = 0;

            while (index < partsListBox.Items.Count)
            {
                if (partsListBox.GetSelected(index))
                {
                    parts.Remove(partsListBox.Items[index].ToString());
                    partsListBox.Items.RemoveAt(index);
                }
                else
                    index++;
            }
        }

        private void UpdateDateLabels(object sender, EventArgs e)
        {
            currentDateLabel.Text = comparisondateDateTimePicker.Value.ToShortDateString();
            iDateLabel.Text = idateDateTimePicker.Value.ToShortDateString();
            date1Label.Text = idateDateTimePicker.Value.ToShortDateString();
            date2Label.Text = comparisondateDateTimePicker.Value.ToShortDateString();
        }

        void dataGridView_iDate_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            object o = dataGridView_iDate.Rows[e.RowIndex].HeaderCell.Value;

            e.Graphics.DrawString(
                o != null ? o.ToString() : "",
                dataGridView_iDate.Font,
                Brushes.Black,
                new PointF((float)e.RowBounds.Left + 2, (float)e.RowBounds.Top + 4));
        }

        void quarter_iDateGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            object o = monthly_iDateGridView.Rows[e.RowIndex].HeaderCell.Value;

            e.Graphics.DrawString(
                o != null ? o.ToString() : "",
                monthly_iDateGridView.Font,
                Brushes.Black,
                new PointF((float)e.RowBounds.Left + 2, (float)e.RowBounds.Top + 4));
        }

        void quarter_currentGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            object o = monthly_currentGridView.Rows[e.RowIndex].HeaderCell.Value;

            e.Graphics.DrawString(
                o != null ? o.ToString() : "",
                monthly_currentGridView.Font,
                Brushes.Black,
                new PointF((float)e.RowBounds.Left + 2, (float)e.RowBounds.Top + 4));
        }

        void dataGridView_current_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            object o = dataGridView_current.Rows[e.RowIndex].HeaderCell.Value;

            e.Graphics.DrawString(
                o != null ? o.ToString() : "",
                dataGridView_current.Font,
                Brushes.Black,
                new PointF((float)e.RowBounds.Left + 2, (float)e.RowBounds.Top + 4));
        }

        private void updateLabels(object sender, DataGridViewCellEventArgs e)
        {
            //
            // Sets up the Labels for INFO
            //
            int rowIndex = linesDataGridView.SelectedRows[0].Index;
            // set part label
            partLabel.Text = linesDataGridView.Rows[rowIndex].Cells[3].Value.ToString();
            // set rev label
            partRevLabel.Text = linesDataGridView.Rows[rowIndex].Cells[12].Value.ToString();
            // set description label
            descriptionLabel.Text = linesDataGridView.Rows[rowIndex].Cells[13].Value.ToString();
            // set customer label
            customerLabel.Text = linesDataGridView.Rows[rowIndex].Cells[2].Value.ToString();
        }

        private void Update_30d90d6m1yTable(object sender, EventArgs e)
        {
            if (linesDataGridView.Rows.Count == 0)
                return;
            //
            // Queries delivery info
            //

            // Fonts that will be used
            DataGridViewCellStyle style = new DataGridViewCellStyle();
            style.ForeColor = Color.Red;

            string filter = string.Empty;
            if (partsRadioButton.Checked)
                filter = "Material = '" + linesDataGridView.SelectedRows[0].Cells[3].Value.ToString() + "'";
            else if (customersRadioButton.Checked)
                filter = "Customer = '" + linesDataGridView.SelectedRows[0].Cells[2].Value.ToString() + "'";
            else if (ByAllRadioButton.Checked)
                filter = "1 = 1";
            else
                filter = "Material = '" + linesDataGridView.SelectedRows[0].Cells[3].Value.ToString() + "' AND Customer = '" + linesDataGridView.SelectedRows[0].Cells[2].Value.ToString() + "'";

            using (OdbcConnection con = new OdbcConnection("DSN=" + DSN + ";UID=" + userName + ";PWD=" + password))
            {
                // Open by iDate
                string query =
                    "SELECT SUM(Order_Qty)\n" +
                    "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                    "WHERE Status NOT IN ('Removed', 'Shipped', 'Closed') AND SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) =\n" +
                        "\t(SELECT CONVERT(DATETIME, MAX(InnerT.idate))\n" +
                        "\tFROM ATIDelivery.dbo.Delivery_Lines AS InnerT\n" +
                        "\tWHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + idateDateTimePicker.Value.ToShortDateString() + "'));";

                con.Open();

                OdbcCommand com = new OdbcCommand(query, con);
                OdbcDataReader reader = com.ExecuteReader();
                reader.Read();
                piecesDue_idateLabel.Text = reader.IsDBNull(0) ? "0" : reader.GetDouble(0).ToString();

                // Open total at comparison date time picker
                query =
                    "SELECT SUM(Order_Qty)\n" +
                    "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                    "WHERE Status NOT IN ('Removed', 'Shipped', 'Closed') AND SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) = " +
                        "(SELECT CONVERT(DATETIME, MAX(InnerT.idate)) " +
                        "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT " +
                        "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + comparisondateDateTimePicker.Value.ToShortDateString() + "'));";


                com = new OdbcCommand(query, con);
                reader = com.ExecuteReader();
                reader.Read();
                piecesDue_currentLabel.Text = reader.IsDBNull(0) ? "0" : reader.GetDouble(0).ToString();

                //
                //==============================DATE 1===========================
                //
                DateTime[] keyDates_picker1 = { idateDateTimePicker.Value, idateDateTimePicker.Value.AddDays(30), idateDateTimePicker.Value.AddDays(90), idateDateTimePicker.Value.AddMonths(6), idateDateTimePicker.Value.AddYears(1) };
                string[] statuses = { "('Open')", "('Hold')" };

                                //
                // Calculate firm and forecast counts for datagridview #1
                //
                string selectQuery = string.Empty;
                for (int i = 0; i < keyDates_picker1.Length; i++)
                    for (int j = 0; j < statuses.Length; j++)
                    {
                        // PD field does not include the current date aka '<' NOT '<='
                        if (i == 0)
                        {
                            selectQuery += "\tSUM(CASE WHEN (Status IN " + statuses[j] + " AND Promised_Date < '" + keyDates_picker1[i].ToShortDateString() + "') THEN Order_Qty ELSE 0 END),\n";
                            selectQuery += "\tSUM(CASE WHEN (Status IN " + statuses[j] + " AND Promised_Date < '" + keyDates_picker1[i].ToShortDateString() + "') THEN Total_Price ELSE 0 END)";
                        }
                        else
                        {
                            selectQuery += "\tSUM(CASE WHEN (Status IN " + statuses[j] + " AND Promised_Date <= '" + keyDates_picker1[i].ToShortDateString() + "') THEN Order_Qty ELSE 0 END),\n";
                            selectQuery += "\tSUM(CASE WHEN (Status IN " + statuses[j] + " AND Promised_Date <= '" + keyDates_picker1[i].ToShortDateString() + "') THEN Total_Price ELSE 0 END)";
                        }
                        // make sure that last row doesn't have an extra comma
                        if (j != statuses.Length - 1 || i != keyDates_picker1.Length - 1)
                            selectQuery += ",\n";
                        else
                            selectQuery += "\n";
                    }

                query =
                    "SELECT\n" + selectQuery +
                                    "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                                    "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) =\n" +
                                        "\t(SELECT CONVERT(DATETIME, MAX(InnerT.idate))\n" +
                                        "\tFROM ATIDelivery.dbo.Delivery_Lines AS InnerT\n" +
                                        "\tWHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + idateDateTimePicker.Value.ToShortDateString() + "'));\n";

                // the costs for pd, 30d, 6m, 1y will be stored in this variable
                // this cost makes use of the filters that the monthly tables use to calculate cost
                decimal[] totalCosts = new decimal[5];
                // initialize to 0
                for (int i = 0; i < totalCosts.Length; i++)
                    totalCosts[i] = 0;
                Boolean[] moneyFilters = { firmCostCheckBox.Checked, forecastCostCheckBox.Checked, shippedCostCheckBox.Checked }; // this filters tell you what costs to include

                com = new OdbcCommand(query, con);
                reader = com.ExecuteReader();
                reader.Read();
                double orderQty;

                for (int status_index = 0; status_index < 2; status_index++)
                {
                    for (int date_index = 0; date_index < keyDates_picker1.Length; date_index++)
                    {
                        orderQty = reader.IsDBNull((date_index * 2 + status_index) * 2) ? 0 : reader.GetDouble((date_index * 2 + status_index) * 2);
                        dataGridView_iDate.Rows[date_index].Cells[status_index].Value = orderQty;
                        totalCosts[date_index] += !reader.IsDBNull((date_index * 2 + status_index) * 2 + 1) && moneyFilters[status_index] ? reader.GetDecimal((date_index * 2 + status_index) * 2 + 1) : 0;
                        if (orderQty > 0 && date_index == 0)
                            dataGridView_iDate.Rows[date_index].Cells[status_index].Style = style;
                        else
                            dataGridView_iDate.Rows[date_index].Cells[status_index].Style = dataGridView_iDate.Columns[status_index].DefaultCellStyle;
                    }
                }

                //
                // Calculate shipping cost for datagridview #1 
                //
                DateTime[] shipping_keyDates_picker1 = { idateDateTimePicker.Value, idateDateTimePicker.Value.AddDays(-30), idateDateTimePicker.Value.AddDays(-90), idateDateTimePicker.Value.AddMonths(-6), idateDateTimePicker.Value.AddYears(-1) };
                selectQuery = string.Empty;

                for (int i = 0; i < keyDates_picker1.Length; i++)
                {
                    selectQuery += "\tSUM(CASE WHEN (Status IN ('Shipped') AND idate >= '" + shipping_keyDates_picker1[i].ToShortDateString() + "') THEN Order_Qty ELSE 0 END),\n";
                    selectQuery += "\tSUM(CASE WHEN (Status IN ('Shipped') AND idate >= '" + shipping_keyDates_picker1[i].ToShortDateString() + "') THEN Total_Price ELSE 0 END)";
                    if (i != shipping_keyDates_picker1.Length - 1)
                        selectQuery += ",\n";
                    else
                        selectQuery += "\n";
                }

                query =
                    "SELECT " + selectQuery + " " +
                            "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                            "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) =\n" +
                                "(SELECT CONVERT(DATETIME, MAX(InnerT.idate))\n" +
                                "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT\n" +
                                "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND InnerT.Status = 'Shipped' AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + idateDateTimePicker.Value.ToShortDateString() + "'));";

                com = new OdbcCommand(query, con);
                reader = com.ExecuteReader();
                reader.Read();
                for (int i = 0; i < shipping_keyDates_picker1.Length; i++)
                {
                    orderQty = reader.IsDBNull(i * 2) ? 0 : reader.GetDouble(i * 2);
                    // TO_DO: Reivew line
                    totalCosts[i] += !reader.IsDBNull(i * 2 + 1) && moneyFilters[2] ? reader.GetDecimal(i * 2 + 1) : 0;
                    dataGridView_iDate.Rows[i].Cells[2].Value = orderQty;
                }

                // iDate Total pd
                double total = double.Parse(dataGridView_iDate.Rows[0].Cells[0].Value.ToString()) + double.Parse(dataGridView_iDate.Rows[0].Cells[1].Value.ToString());
                dataGridView_iDate.Rows[0].Cells[3].Value = total;
                if (total > 0)
                    dataGridView_iDate.Rows[0].Cells[3].Style = style;
                else
                    dataGridView_iDate.Rows[0].Cells[3].Style = dataGridView_iDate.Columns[2].DefaultCellStyle;

                // iDate Total 30d
                dataGridView_iDate.Rows[1].Cells[3].Value = double.Parse(dataGridView_iDate.Rows[1].Cells[0].Value.ToString()) + double.Parse(dataGridView_iDate.Rows[1].Cells[1].Value.ToString());

                // iDate Total 90d
                dataGridView_iDate.Rows[2].Cells[3].Value = double.Parse(dataGridView_iDate.Rows[2].Cells[0].Value.ToString()) + double.Parse(dataGridView_iDate.Rows[2].Cells[1].Value.ToString());

                // iDate Total 6 mo
                dataGridView_iDate.Rows[3].Cells[3].Value = double.Parse(dataGridView_iDate.Rows[3].Cells[0].Value.ToString()) + double.Parse(dataGridView_iDate.Rows[3].Cells[1].Value.ToString());

                // iDate Total 1yr
                dataGridView_iDate.Rows[4].Cells[3].Value = double.Parse(dataGridView_iDate.Rows[4].Cells[0].Value.ToString()) + double.Parse(dataGridView_iDate.Rows[4].Cells[1].Value.ToString());


                // write total cost data to datagridview
                for (int i = 0; i < totalCosts.Length; i++)
                    dataGridView_iDate.Rows[i].Cells[4].Value = totalCosts[i];

                dataGridView_iDate.Columns[4].DefaultCellStyle.Format = "c";

                //
                //==============================DATE 2===========================
                //

                // Date2 Firm pd
                DateTime[] keyDates_picker2 = { comparisondateDateTimePicker.Value, comparisondateDateTimePicker.Value.AddDays(30), comparisondateDateTimePicker.Value.AddDays(90), comparisondateDateTimePicker.Value.AddMonths(6), comparisondateDateTimePicker.Value.AddYears(1) };

                //
                // Calculate firm and forecast counts for datagridview #1
                //
                selectQuery = string.Empty;
                for (int i = 0; i < keyDates_picker2.Length; i++)
                    for (int j = 0; j < statuses.Length; j++)
                    {
                        // PD field does not include the current date aka '<' NOT '<='
                        if (i == 0)
                        {
                            selectQuery += "\tSUM(CASE WHEN (Status IN " + statuses[j] + " AND Promised_Date < '" + keyDates_picker2[i].ToShortDateString() + "') THEN Order_Qty ELSE 0 END),\n";
                            selectQuery += "\tSUM(CASE WHEN (Status IN " + statuses[j] + " AND Promised_Date < '" + keyDates_picker2[i].ToShortDateString() + "') THEN Total_Price ELSE 0 END)";
                        }
                        else
                        {
                            selectQuery += "\tSUM(CASE WHEN (Status IN " + statuses[j] + " AND Promised_Date <= '" + keyDates_picker2[i].ToShortDateString() + "') THEN Order_Qty ELSE 0 END),\n";
                            selectQuery += "\tSUM(CASE WHEN (Status IN " + statuses[j] + " AND Promised_Date <= '" + keyDates_picker2[i].ToShortDateString() + "') THEN Total_Price ELSE 0 END)";
                        }
                        // make sure that last row doesn't have an extra comma
                        if (j != statuses.Length - 1 || i != keyDates_picker2.Length - 1)
                            selectQuery += ",\n";
                        else
                            selectQuery += "\n";
                    }

                query =
                    "SELECT\n" + selectQuery +
                                    "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                                    "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) =\n" +
                                        "\t(SELECT CONVERT(DATETIME, MAX(InnerT.idate))\n" +
                                        "\tFROM ATIDelivery.dbo.Delivery_Lines AS InnerT\n" +
                                        "\tWHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + comparisondateDateTimePicker.Value.ToShortDateString() + "'));\n";

                // the costs for pd, 30d, 6m, 1y will be stored in this variable
                // this cost makes use of the filters that the monthly tables use to calculate cost
                // initialize to 0
                for (int i = 0; i < totalCosts.Length; i++)
                    totalCosts[i] = 0;

                com = new OdbcCommand(query, con);
                reader = com.ExecuteReader();
                reader.Read();
                orderQty = 0;

                for (int status_index = 0; status_index < 2; status_index++)
                {
                    for (int date_index = 0; date_index < keyDates_picker2.Length; date_index++)
                    {
                        orderQty = reader.IsDBNull((date_index * 2 + status_index) * 2) ? 0 : reader.GetDouble((date_index * 2 + status_index) * 2);
                        dataGridView_current.Rows[date_index].Cells[status_index].Value = orderQty;
                        totalCosts[date_index] += !reader.IsDBNull((date_index * 2 + status_index) * 2 + 1) && moneyFilters[status_index] ? reader.GetDecimal((date_index * 2 + status_index) * 2 + 1) : 0;
                        if (orderQty > 0 && date_index == 0)
                            dataGridView_current.Rows[date_index].Cells[status_index].Style = style;
                        else
                            dataGridView_current.Rows[date_index].Cells[status_index].Style = dataGridView_current.Columns[status_index].DefaultCellStyle;
                    }
                }

                //
                // Calculate shipping cost for datagridview #1 
                //
                DateTime[] shipping_keyDates_picker2 = { comparisondateDateTimePicker.Value, comparisondateDateTimePicker.Value.AddDays(-30), comparisondateDateTimePicker.Value.AddDays(-90), comparisondateDateTimePicker.Value.AddMonths(-6), comparisondateDateTimePicker.Value.AddYears(-1) };
                selectQuery = string.Empty;

                for (int i = 0; i < keyDates_picker2.Length; i++)
                {
                    selectQuery += "\tSUM(CASE WHEN (Status IN ('Shipped') AND idate >= '" + shipping_keyDates_picker2[i].ToShortDateString() + "') THEN Order_Qty ELSE 0 END),\n";
                    selectQuery += "\tSUM(CASE WHEN (Status IN ('Shipped') AND idate >= '" + shipping_keyDates_picker2[i].ToShortDateString() + "') THEN Total_Price ELSE 0 END)";
                    if (i != shipping_keyDates_picker2.Length - 1)
                        selectQuery += ",\n";
                    else
                        selectQuery += "\n";
                }

                query =
                    "SELECT " + selectQuery + " " +
                            "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                            "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) =\n" +
                                "(SELECT CONVERT(DATETIME, MAX(InnerT.idate))\n" +
                                "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT\n" +
                                "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND InnerT.Status = 'Shipped' AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + comparisondateDateTimePicker.Value.ToShortDateString() + "'));";

                com = new OdbcCommand(query, con);
                reader = com.ExecuteReader();
                reader.Read();
                for (int i = 0; i < shipping_keyDates_picker2.Length; i++)
                {
                    orderQty = reader.IsDBNull(i * 2) ? 0 : reader.GetDouble(i * 2);
                    // TO_DO: Reivew line
                    totalCosts[i] += !reader.IsDBNull(i * 2 + 1) && moneyFilters[2] ? reader.GetDecimal(i * 2 + 1) : 0;
                    dataGridView_current.Rows[i].Cells[2].Value = orderQty;
                }

                // iDate Total pd
                total = double.Parse(dataGridView_current.Rows[0].Cells[0].Value.ToString()) + double.Parse(dataGridView_current.Rows[0].Cells[1].Value.ToString());
                dataGridView_current.Rows[0].Cells[3].Value = total;
                if (total > 0)
                    dataGridView_current.Rows[0].Cells[3].Style = style;
                else
                    dataGridView_current.Rows[0].Cells[3].Style = dataGridView_current.Columns[2].DefaultCellStyle;

                // iDate Total 30d
                dataGridView_current.Rows[1].Cells[3].Value = double.Parse(dataGridView_current.Rows[1].Cells[0].Value.ToString()) + double.Parse(dataGridView_current.Rows[1].Cells[1].Value.ToString());

                // iDate Total 90d
                dataGridView_current.Rows[2].Cells[3].Value = double.Parse(dataGridView_current.Rows[2].Cells[0].Value.ToString()) + double.Parse(dataGridView_current.Rows[2].Cells[1].Value.ToString());

                // iDate Total 6 mo
                dataGridView_current.Rows[3].Cells[3].Value = double.Parse(dataGridView_current.Rows[3].Cells[0].Value.ToString()) + double.Parse(dataGridView_current.Rows[3].Cells[1].Value.ToString());

                // iDate Total 1yr
                dataGridView_current.Rows[4].Cells[3].Value = double.Parse(dataGridView_current.Rows[4].Cells[0].Value.ToString()) + double.Parse(dataGridView_current.Rows[4].Cells[1].Value.ToString());


                // write total cost data to datagridview
                for (int i = 0; i < totalCosts.Length; i++)
                    dataGridView_current.Rows[i].Cells[4].Value = totalCosts[i];

                dataGridView_current.Columns[4].DefaultCellStyle.Format = "c";

                //
                // Calculate total ship between the two dates
                //

                if (idateDateTimePicker.Value < comparisondateDateTimePicker.Value)
                {

                    query = "SELECT SUM(CASE WHEN (Status IN ('Shipped') AND CONVERT(DATE,idate) <= CONVERT(DATE, '" + comparisondateDateTimePicker.Value.ToShortDateString() + "') AND CONVERT(DATE, idate) >= CONVERT(DATE,'" + idateDateTimePicker.Value.ToShortDateString() + "')) THEN Order_Qty ELSE 0 END)\n" +
                            "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                            "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND Material = 'CH2151-0016' AND CONVERT(DATETIME, idate) =\n" +
                                "\t(SELECT CONVERT(DATETIME, MAX(InnerT.idate))\n" +
                                "\tFROM ATIDelivery.dbo.Delivery_Lines AS InnerT\n" +
                                "\tWHERE InnerT.SO_Detail = OutterT.SO_Detail AND InnerT.Status = 'Shipped' AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + comparisondateDateTimePicker.Value.ToShortDateString() + "'));";

                    com = new OdbcCommand(query, con);
                    reader = com.ExecuteReader();

                    reader.Read();

                    switch (Type.GetTypeCode(reader.GetFieldType(0)))
                    {
                        case TypeCode.Int32:
                            partsShippedLabel.Text = reader.IsDBNull(0) ? (0).ToString() : reader.GetInt32(0).ToString();
                            break;
                        case TypeCode.Double:
                            partsShippedLabel.Text = reader.IsDBNull(0) ? (0).ToString() : reader.GetDouble(0).ToString();
                            break;
                    }
                }
                else
                {
                    partsShippedLabel.Text = "0";
                }
            }
        }

        private void Update_Monthly(object sender, EventArgs e)
        {
            // check if there is any lines to pull information off of
            if (iDate_yearDropDown.SelectedItem == null || linesDataGridView.Rows.Count == 0)
                return;

            // disable all the buttons to avoid the user from stagnating clicks
            linesDataGridView.Enabled = false;
            idateDateTimePicker.Enabled = false;
            iDate_yearDropDown.Enabled = false;
            changesRadioButton.Enabled = false;
            openLinesRadioButton.Enabled = false;
            partsRadioButton.Enabled = false;
            customersRadioButton.Enabled = false;
            customerAndPartsRadioButton.Enabled = false;
            ByAllRadioButton.Enabled = false;
            addButton.Enabled = false;
            firmCostCheckBox.Enabled = false;
            shippedCostCheckBox.Enabled = false;
            forecastCostCheckBox.Enabled = false;
            selectAllCostCheckBox.Enabled = false;

            // get the part number of the selected row
            string part = linesDataGridView.SelectedRows[0].Cells[3].Value.ToString();
            // get the selected year
            // this will decide what set of months we'll be working with
            int year = int.Parse(iDate_yearDropDown.SelectedItem.ToString());

            // the three possible statuses
            /*
             * Open -> Firm
             * Hold -> Forecast
             * Shipped -> Shipped
             */
            string[] statuses = { "('Open')", "('Hold')", "('Shipped')" };
            Boolean[] moneyFilters = { firmCostCheckBox.Checked, forecastCostCheckBox.Checked, shippedCostCheckBox.Checked };

            // this sets the part number and customer filter vased on the selections of the "Filter By" dropdown
            string filter = string.Empty;
            if (partsRadioButton.Checked)
                filter = "Material = '" + linesDataGridView.SelectedRows[0].Cells[3].Value.ToString() + "'";
            else if (customersRadioButton.Checked)
                filter = "Customer = '" + linesDataGridView.SelectedRows[0].Cells[2].Value.ToString() + "'";
            else if (ByAllRadioButton.Checked)
                filter = "1 = 1";
            else
                filter = "Material = '" + linesDataGridView.SelectedRows[0].Cells[3].Value.ToString() + "' AND Customer = '" + linesDataGridView.SelectedRows[0].Cells[2].Value.ToString() + "'";

            //
            // IDATE - left data gridview of the monthly set
            //

            // this will be used to calculate the cost of the 12 months
            decimal[] totalCosts = new decimal[12];

            using (OdbcConnection con = new OdbcConnection("DSN=" + DSN + ";UID=" + userName + ";PWD=" + password))
            {
                con.Open();
                string selectString_Forecast = string.Empty;
                string selectString_Shipping = string.Empty;
                string selectString_totalDue = string.Empty;
                string dateRangeVar = string.Empty;

                for (int j = 0; j < 2; j++)
                    for (int i = 1; i <= 12; i++)
                    {
                        dateRangeVar = "Promised_Date";
                        selectString_Forecast += "\tSUM(CASE WHEN (Status IN " + statuses[j] + " AND CONVERT(DATE," + dateRangeVar + ") <= CONVERT(DATE, '" + (new DateTime(year, i, DateTime.DaysInMonth(year, i))).ToShortDateString() + "') AND CONVERT(DATE, " + dateRangeVar + ") >= CONVERT(DATE, '" + (new DateTime(year, i, 1)).ToShortDateString() + "')) THEN Order_Qty ELSE 0 END),\n" +
                                                 "\tSUM(CASE WHEN (Status IN " + statuses[j] + " AND CONVERT(DATE, " + dateRangeVar + ") <= CONVERT(DATE, '" + (new DateTime(year, i, DateTime.DaysInMonth(year, i))).ToShortDateString() + "') AND CONVERT(DATE, " + dateRangeVar + ") >= CONVERT(DATE, '" + (new DateTime(year, i, 1)).ToShortDateString() + "')) THEN Total_Price ELSE 0 END)";
                        if (!(j == 1 && i == 12))
                            selectString_Forecast += ",\n";
                        else
                            selectString_Forecast += "\n";
                    }

                string query =
                    "SELECT\n" + selectString_Forecast +
                            "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                            "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) =\n" +
                                "(SELECT CONVERT(DATETIME, MAX(InnerT.idate))\n" +
                                "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT\n" +
                                "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + idateDateTimePicker.Value.ToShortDateString() + "'));\n";
                OdbcCommand com = new OdbcCommand(query, con);
                OdbcDataReader reader = com.ExecuteReader();

                reader.Read();
                double orderQty;
                for (int j = 0; j < 2; j++)
                    for (int i = 0; i < 12; i++)
                    {
                        orderQty = reader.IsDBNull((j * 12 + i) * 2) ? 0 : reader.GetDouble((j * 12 + i) * 2);
                        monthly_iDateGridView.Rows[i].Cells[j].Value = orderQty;
                        totalCosts[i] += !reader.IsDBNull((j * 12 + i) * 2 + 1) && moneyFilters[j] ? reader.GetDecimal((j * 12 + i) * 2 + 1) : 0;
                    }
                reader.Read();

                //==============SHIPPING==========================

                for (int j = 2; j < 3; j++)
                    for (int i = 1; i <= 12; i++)
                    {
                        dateRangeVar = "idate";

                        selectString_Shipping += "SUM(CASE WHEN (Status IN " + statuses[j] + " AND CONVERT(DATE," + dateRangeVar + ") <= CONVERT(DATE, '" + (new DateTime(year, i, DateTime.DaysInMonth(year, i))).ToShortDateString() + "') AND CONVERT(DATE, " + dateRangeVar + ") >= CONVERT(DATE,'" + (new DateTime(year, i, 1)).ToShortDateString() + "')) THEN Order_Qty ELSE 0 END),\n" +
                                    "SUM(CASE WHEN (Status IN " + statuses[j] + " AND CONVERT(DATE," + dateRangeVar + ") <= CONVERT(DATE,'" + (new DateTime(year, i, DateTime.DaysInMonth(year, i))).ToShortDateString() + "') AND CONVERT(DATE," + dateRangeVar + ") >= CONVERT(DATE,'" + (new DateTime(year, i, 1)).ToShortDateString() + "')) THEN Total_Price ELSE 0 END)";
                        if (!(j == 2 && i == 12))
                            selectString_Shipping += ",\n";
                        else
                            selectString_Shipping += "\n";
                    }

                query =
                    "SELECT " + selectString_Shipping + " " +
                            "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                            "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) =\n" +
                                "(SELECT CONVERT(DATETIME, MAX(InnerT.idate))\n" +
                                "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT\n" +
                                "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND InnerT.Status = 'Shipped' AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + idateDateTimePicker.Value.ToShortDateString() + "'));";
                com = new OdbcCommand(query, con);
                reader = com.ExecuteReader();

                reader.Read();
                for (int j = 2; j < 3; j++)
                    for (int i = 0; i < 12; i++)
                    {
                        orderQty = reader.IsDBNull(i * 2) ? 0 : reader.GetDouble(i * 2);
                        monthly_iDateGridView.Rows[i].Cells[j].Value = orderQty;
                        totalCosts[i] += !reader.IsDBNull(i * 2 + 1) && moneyFilters[j] ? reader.GetDecimal(i * 2 + 1) : 0;
                    }

                reader.Read();

                // =========================================

                for (int i = 1; i <= 12; i++)
                {
                    selectString_totalDue += "SUM(CASE WHEN (Promised_Date >= CONVERT(DATE, '" + (new DateTime(year, i, 1)).ToShortDateString() + "') AND Promised_Date <= CONVERT(DATE, '" + (new DateTime(year, i, DateTime.DaysInMonth(year, i))).ToShortDateString() + "')) THEN Order_Qty ELSE 0 END)";
                    if (i != 12)
                        selectString_totalDue += ",\n";
                    else
                        selectString_totalDue += "\n";
                }

                query = "SELECT " + selectString_totalDue + " " +
                        "FROM ATIDelivery.dbo.Delivery_Lines AS outT\n" +
                        "WHERE " + filter + " AND Status <> 'Removed' AND SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND CONVERT(DATETIME, idate) =\n" +
                            "\t(SELECT CONVERT(DATETIME, MAX(inTable.idate))\n" +
                            "\tFROM ATIDelivery.dbo.Delivery_Lines AS inTable\n" +
                            "\tWHERE outT.SO_Detail = inTable.SO_Detail AND CONVERT(DATE, idate) <= CONVERT(DATE, '" + idateDateTimePicker.Value.ToShortDateString() + "')); ";
                com = new OdbcCommand(query, con);
                reader = com.ExecuteReader();

                reader.Read();
                for (int i = 0; i < 12; i++)
                {
                    monthly_iDateGridView.Rows[i].Cells[4].Value = totalCosts[i];
                    monthly_iDateGridView.Rows[i].Cells[3].Value = reader.IsDBNull(i) ? 0 : reader.GetDouble(i);
                }
                reader.Read();
                monthly_iDateGridView.Columns[4].DefaultCellStyle.Format = "c";

                //===================================================================================================================================
                //===================================================================================================================================

                //
                // CURRENT DATE
                //
                totalCosts = new decimal[12];

                query =
                    "SELECT " + selectString_Forecast + " " +
                            "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT " +
                            "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) = " +
                                                "(SELECT CONVERT(DATETIME, MAX(InnerT.idate)) " +
                                                "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT " +
                                                "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + comparisondateDateTimePicker.Value.ToShortDateString() + "'));";


                com = new OdbcCommand(query, con);
                reader = com.ExecuteReader();
                reader.Read();
                for (int j = 0; j < 2; j++)
                    for (int i = 0; i < 12; i++)
                    {
                        orderQty = reader.IsDBNull((j * 12 + i) * 2) ? 0 : reader.GetDouble((j * 12 + i) * 2);
                        monthly_currentGridView.Rows[i].Cells[j].Value = orderQty;
                        totalCosts[i] += !reader.IsDBNull((j * 12 + i) * 2 + 1) && moneyFilters[j] ? reader.GetDecimal((j * 12 + i) * 2 + 1) : 0;
                    }
                reader.Read();


                // =========================================================

                query =
                    "SELECT " + selectString_Shipping + " " +
                            "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT " +
                            "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) = " +
                                                "(SELECT CONVERT(DATETIME, MAX(InnerT.idate)) " +
                                                "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT " +
                                                "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND InnerT.Status = 'Shipped' AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + comparisondateDateTimePicker.Value.ToShortDateString() + "'));";


                com = new OdbcCommand(query, con);
                reader = com.ExecuteReader();
                reader.Read();
                for (int j = 2; j < 3; j++)
                    for (int i = 0; i < 12; i++)
                    {
                        orderQty = reader.IsDBNull(i * 2) ? 0 : reader.GetDouble(i * 2);
                        monthly_currentGridView.Rows[i].Cells[j].Value = orderQty;
                        totalCosts[i] += !reader.IsDBNull(i * 2 + 1) && moneyFilters[j] ? reader.GetDecimal(i * 2 + 1) : 0;
                    }
                reader.Read();

                //==============================================================================

                query = "SELECT " + selectString_totalDue + " " +
                        "FROM ATIDelivery.dbo.Delivery_Lines AS outT\n" +
                        "WHERE " + filter + " AND Status <> 'Removed' AND SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND  CONVERT(DATETIME, idate) =\n" +
                            "\t(SELECT CONVERT(DATETIME, MAX(inTable.idate))\n" +
                            "\tFROM ATIDelivery.dbo.Delivery_Lines AS inTable\n" +
                            "\tWHERE outT.SO_Detail = inTable.SO_Detail AND CONVERT(DATETIME, idate) <= CONVERT(DATE, '" + comparisondateDateTimePicker.Value.ToShortDateString() + "')); ";
                com = new OdbcCommand(query, con);
                reader = com.ExecuteReader();

                reader.Read();
                for (int i = 0; i < 12; i++)
                {
                    monthly_currentGridView.Rows[i].Cells[4].Value = totalCosts[i];
                    monthly_currentGridView.Rows[i].Cells[3].Value = reader.IsDBNull(i) ? 0 : reader.GetDouble(i); ;
                }
                monthly_currentGridView.Columns[4].DefaultCellStyle.Format = "c";
                reader.Read();
            }
            // ==========================================================


            linesDataGridView.Enabled = true;
            idateDateTimePicker.Enabled = true;
            iDate_yearDropDown.Enabled = true;
            changesRadioButton.Enabled = true;
            openLinesRadioButton.Enabled = true;
            partsRadioButton.Enabled = true;
            customersRadioButton.Enabled = true;
            customerAndPartsRadioButton.Enabled = true;
            ByAllRadioButton.Enabled = true;
            addButton.Enabled = true;
            firmCostCheckBox.Enabled = true;
            shippedCostCheckBox.Enabled = true;
            forecastCostCheckBox.Enabled = true;
            selectAllCostCheckBox.Enabled = true;

        }

        private void selectAllCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.firmCostCheckBox.CheckedChanged -= this.firmCheckBox_CheckedChanged;
            this.firmCostCheckBox.CheckedChanged -= this.Update_Monthly;
            this.firmCostCheckBox.CheckedChanged -= this.Update_30d90d6m1yTable;
            this.forecastCostCheckBox.CheckedChanged -= this.forecastCheckBox_CheckedChanged;
            this.forecastCostCheckBox.CheckedChanged -= this.Update_Monthly;
            this.forecastCostCheckBox.CheckedChanged -= this.Update_30d90d6m1yTable;
            this.shippedCostCheckBox.CheckedChanged -= this.totalCheckBox_CheckedChanged;
            this.shippedCostCheckBox.CheckedChanged -= this.Update_Monthly;
            this.shippedCostCheckBox.CheckedChanged -= this.Update_30d90d6m1yTable;
            firmCostCheckBox.Checked = true;
            forecastCostCheckBox.Checked = true;
            shippedCostCheckBox.Checked = true;
            this.firmCostCheckBox.CheckedChanged += this.firmCheckBox_CheckedChanged;
            this.firmCostCheckBox.CheckedChanged += this.Update_Monthly;
            this.firmCostCheckBox.CheckedChanged += this.Update_30d90d6m1yTable;
            this.forecastCostCheckBox.CheckedChanged += this.forecastCheckBox_CheckedChanged;
            this.forecastCostCheckBox.CheckedChanged += this.Update_Monthly;
            this.forecastCostCheckBox.CheckedChanged += this.Update_30d90d6m1yTable;
            this.shippedCostCheckBox.CheckedChanged += this.totalCheckBox_CheckedChanged;
            this.shippedCostCheckBox.CheckedChanged += this.Update_Monthly;
            this.shippedCostCheckBox.CheckedChanged += this.Update_30d90d6m1yTable;
        }

        private void firmCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.selectAllCostCheckBox.CheckedChanged -= this.selectAllCheckBox_CheckedChanged;
            this.selectAllCostCheckBox.CheckedChanged -= this.Update_Monthly;
            this.selectAllCostCheckBox.CheckedChanged -= this.Update_30d90d6m1yTable;
            selectAllCostCheckBox.Checked = false;
            this.selectAllCostCheckBox.CheckedChanged += this.selectAllCheckBox_CheckedChanged;
            this.selectAllCostCheckBox.CheckedChanged += this.Update_Monthly;
            this.selectAllCostCheckBox.CheckedChanged += this.Update_30d90d6m1yTable;
        }

        private void forecastCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.selectAllCostCheckBox.CheckedChanged -= this.selectAllCheckBox_CheckedChanged;
            this.selectAllCostCheckBox.CheckedChanged -= this.Update_Monthly;
            this.selectAllCostCheckBox.CheckedChanged -= this.Update_30d90d6m1yTable;
            selectAllCostCheckBox.Checked = false;
            this.selectAllCostCheckBox.CheckedChanged += this.selectAllCheckBox_CheckedChanged;
            this.selectAllCostCheckBox.CheckedChanged += this.Update_Monthly;
            this.selectAllCostCheckBox.CheckedChanged += this.Update_30d90d6m1yTable;
        }

        private void totalCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.selectAllCostCheckBox.CheckedChanged -= this.selectAllCheckBox_CheckedChanged;
            this.selectAllCostCheckBox.CheckedChanged -= this.Update_Monthly;
            this.selectAllCostCheckBox.CheckedChanged -= this.Update_30d90d6m1yTable;
            selectAllCostCheckBox.Checked = false;
            this.selectAllCostCheckBox.CheckedChanged += this.selectAllCheckBox_CheckedChanged;
            this.selectAllCostCheckBox.CheckedChanged += this.Update_Monthly;
            this.selectAllCostCheckBox.CheckedChanged += this.Update_30d90d6m1yTable;
        }

        private void dgvSomeDataGridView_SelectionChanged(Object sender, EventArgs e)
        {
            ((DataGridView)sender).ClearSelection();
        }

        private void Form1_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            System.Diagnostics.Process.Start(@"T:\Delivery Assistant\SO_Delivery_Interface\SO_Delivery_Interface\bin\Release\helpFile.pdf");
        }
    }
}

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
            startDateTimePicker.Value = DateTime.Today.AddDays(-1);
            startDateTimePicker.MinDate = new DateTime(2015, 7, 29);
            startDateTimePicker.CloseUp += this.UpdateDateLabels;
            startDateTimePicker.CloseUp += this.Update_Monthly;
            startDateTimePicker.CloseUp += this.Update_DueDate;
            startDateTimePicker.CloseUp += this.UpdateMainTable;

            // set up date time picker2
            endDateTimePicker.Value = DateTime.Today;
            endDateTimePicker.MinDate = new DateTime(2015, 7, 29);
            endDateTimePicker.CloseUp += this.UpdateDateLabels;
            endDateTimePicker.CloseUp += this.Update_Monthly;
            endDateTimePicker.CloseUp += this.Update_DueDate;
            endDateTimePicker.CloseUp += this.UpdateMainTable;

            // default radio button for datagridview
            openLinesRadioButton.Checked = true;
            openLinesRadioButton.CheckedChanged += this.UpdateMainTable;

            // default radio button for due dates
            partsRadioButton.Checked = true;
            partsRadioButton.CheckedChanged += this.Update_DueDate;
            partsRadioButton.CheckedChanged += this.Update_Monthly;

            // Set up datagrid view
            UpdateMainTable(new Object(), new EventArgs());
            linesDataGridView.CellClick += this.Update_Monthly;
            linesDataGridView.CellClick += this.Update_DueDate;
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
            quarter_iDateGridView.Rows.Add(12);
            for (int i = 1; i <= 12; i++)
                quarter_iDateGridView.Rows[i - 1].HeaderCell.Value = UpperCaseFirst(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i).Substring(0, 3));
            quarter_iDateGridView.RowHeadersDefaultCellStyle.Padding = new Padding(linesDataGridView.RowHeadersWidth);
            quarter_iDateGridView.RowPostPaint += new DataGridViewRowPostPaintEventHandler(quarter_iDateGridView_RowPostPaint);

            // set up quarters_current
            quarter_currentGridView.Rows.Add(12);
            for (int i = 1; i <= 12; i++)
                quarter_currentGridView.Rows[i - 1].HeaderCell.Value = UpperCaseFirst(CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(i).Substring(0, 3));
            quarter_currentGridView.RowHeadersDefaultCellStyle.Padding = new Padding(linesDataGridView.RowHeadersWidth);
            quarter_currentGridView.RowPostPaint += new DataGridViewRowPostPaintEventHandler(quarter_currentGridView_RowPostPaint);
            quarter_currentGridView.Rows[DateTime.Now.Month - 1].Selected = true;

            // set up date
            currentDateLabel.Text = endDateTimePicker.Value.ToShortDateString();
            date2Label.Text = endDateTimePicker.Value.ToShortDateString();

            // set up iDate
            iDateLabel.Text = startDateTimePicker.Value.ToShortDateString();
            date1Label.Text = startDateTimePicker.Value.ToShortDateString();

            // set up firm checkbox
            firmCheckBox.CheckedChanged += this.Update_Monthly;

            // set up forecast checkbox
            forecastCheckBox.CheckedChanged += this.Update_Monthly;

            // set up total checkbox
            shippedCheckBox.CheckedChanged += this.Update_Monthly;

            // set up selectAll checkbox
            selectAllCheckBox.CheckedChanged += this.Update_Monthly;

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
                                        "\tWHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + startDateTimePicker.Value.ToShortDateString() + "'))\n" +
                                    "ORDER BY Promised_Date, SO_Line;";
                    con.Open();

                    OdbcCommand com = new OdbcCommand(query, con);
                    try
                    {
                        OdbcDataReader read = com.ExecuteReader();

                        DataSet ds = new DataSet();
                        DataTable dt = new DataTable();
                        ds.Tables.Add(dt);
                        ds.Load(read, LoadOption.PreserveChanges, ds.Tables[0]);

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
                                    "WHERE  idate >= '" + startDateTimePicker.Value.ToShortDateString() + "' AND idate < '" + startDateTimePicker.Value.AddDays(1).ToShortDateString() + "' AND ( " + (customers.Count == 0 ? "1 = 1" : "1= 0") + " OR Customer IN ('" + string.Join("','", customers.ToArray()) + "')) AND ( " + (parts.Count == 0 ? "1 = 1" : "1= 0") + " OR Material IN ('" + string.Join("','", parts.ToArray()) + "'))\n" +
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
                                "\t(SELECT * FROM ATIDelivery.dbo.Delivery_Lines AS compareT2 WHERE compareT2.idate = (SELECT MAX(innerT.idate) FROM ATIDelivery.dbo.Delivery_Lines AS innerT WHERE  innerT.SO_Detail = compareT2.SO_Detail AND innerT.idate < CONVERT(DATE, '" + startDateTimePicker.Value.ToShortDateString() + "'))) AS compareT\n" +
                          "ON compareT.SO_Detail = readT.SO_Detail\n" +
                          "WHERE  readT.idate >= '" + startDateTimePicker.Value.ToShortDateString() + "' AND readT.idate < '" + startDateTimePicker.Value.AddDays(1).ToShortDateString() + "'  AND ( " + (customers.Count == 0 ? "1 = 1" : "1= 0") + " OR readT.Customer IN ('" + string.Join("','", customers.ToArray()) + "')) AND ( " + (parts.Count == 0 ? "1 = 1" : "1= 0") + " OR readT.Material IN ('" + string.Join("','", parts.ToArray()) + "'))\n" +
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
            currentDateLabel.Text = endDateTimePicker.Value.ToShortDateString();
            iDateLabel.Text = startDateTimePicker.Value.ToShortDateString();
            date1Label.Text = startDateTimePicker.Value.ToShortDateString();
            date2Label.Text = endDateTimePicker.Value.ToShortDateString();
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
            object o = quarter_iDateGridView.Rows[e.RowIndex].HeaderCell.Value;

            e.Graphics.DrawString(
                o != null ? o.ToString() : "",
                quarter_iDateGridView.Font,
                Brushes.Black,
                new PointF((float)e.RowBounds.Left + 2, (float)e.RowBounds.Top + 4));
        }

        void quarter_currentGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            object o = quarter_currentGridView.Rows[e.RowIndex].HeaderCell.Value;

            e.Graphics.DrawString(
                o != null ? o.ToString() : "",
                quarter_currentGridView.Font,
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

        private void Update_DueDate(object sender, EventArgs e)
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
                    "SELECT SUM(Order_Qty) " +
                                    "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT " +
                                    "WHERE Status NOT IN ('Removed', 'Shipped', 'Closed') AND SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) = " +
                                        "(SELECT CONVERT(DATETIME, MAX(InnerT.idate)) " +
                                        "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT " +
                                        "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + startDateTimePicker.Value.ToShortDateString() + "'));";

                con.Open();

                OdbcCommand com = new OdbcCommand(query, con);
                OdbcDataReader read = com.ExecuteReader();
                read.Read();
                openTotal_idateLabel.Text = read.IsDBNull(0) ? "0" : read.GetDouble(0).ToString();

                // Open total
                query =
                    "SELECT SUM(Order_Qty) " +
                                    "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT " +
                                    "WHERE Status NOT IN ('Removed', 'Shipped', 'Closed') AND SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) = " +
                                        "(SELECT CONVERT(DATETIME, MAX(InnerT.idate)) " +
                                        "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT " +
                                        "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + endDateTimePicker.Value.ToShortDateString() + "'));";


                com = new OdbcCommand(query, con);
                read = com.ExecuteReader();
                read.Read();
                openTotal_currentLabel.Text = read.IsDBNull(0) ? "0" : read.GetDouble(0).ToString();


                DateTime[] keyDates_picker1 = { startDateTimePicker.Value, startDateTimePicker.Value.AddDays(30), startDateTimePicker.Value.AddDays(90), startDateTimePicker.Value.AddMonths(6), startDateTimePicker.Value.AddYears(1) };
                string[] statuses = { "('Hold')", "('Open')" };


                // START OF REPETITITITITIONS======================================================
                string selectQuery = string.Empty;
                for (int i = 0; i < keyDates_picker1.Length; i++)
                    for (int j = 0; j < statuses.Length; j++)
                    {
                        selectQuery += "SUM(CASE WHEN (Status IN " + statuses[j] + " AND Promised_Date <= '" + keyDates_picker1[i].ToShortDateString() + "') THEN Order_Qty ELSE 0 END)";
                        if (j != statuses.Length - 1 || i != keyDates_picker1.Length - 1)
                            selectQuery += ", ";
                        else
                            selectQuery += " ";
                    }

                query =
                    "SELECT " + selectQuery +
                                    "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT " +
                                    "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) = " +
                                        "(SELECT CONVERT(DATETIME, MAX(InnerT.idate)) " +
                                        "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT " +
                                        "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + startDateTimePicker.Value.ToShortDateString() + "'));";

                com = new OdbcCommand(query, con);
                read = com.ExecuteReader();
                read.Read();
                for (int i = 0; i < keyDates_picker1.Length; i++)
                    for (int j = 0; j < statuses.Length; j++)
                    {
                        double orderQty = read.IsDBNull(i * statuses.Length + j) ? 0 : read.GetDouble(i * statuses.Length + j);
                        dataGridView_iDate.Rows[i].Cells[j].Value = orderQty;
                        if (orderQty > 0 && i == 0)
                            dataGridView_iDate.Rows[i].Cells[j].Style = style;
                        else
                            dataGridView_iDate.Rows[i].Cells[j].Style = dataGridView_iDate.Columns[j].DefaultCellStyle;
                    }

                // iDate Total pd
                double total = double.Parse(dataGridView_iDate.Rows[0].Cells[0].Value.ToString()) + double.Parse(dataGridView_iDate.Rows[0].Cells[1].Value.ToString());
                dataGridView_iDate.Rows[0].Cells[2].Value = total;
                if (total > 0)
                    dataGridView_iDate.Rows[0].Cells[2].Style = style;
                else
                    dataGridView_iDate.Rows[0].Cells[2].Style = dataGridView_iDate.Columns[2].DefaultCellStyle;

                // iDate Total 30d
                dataGridView_iDate.Rows[1].Cells[2].Value = double.Parse(dataGridView_iDate.Rows[1].Cells[0].Value.ToString()) + double.Parse(dataGridView_iDate.Rows[1].Cells[1].Value.ToString());

                // iDate Total 90d
                dataGridView_iDate.Rows[2].Cells[2].Value = double.Parse(dataGridView_iDate.Rows[2].Cells[0].Value.ToString()) + double.Parse(dataGridView_iDate.Rows[2].Cells[1].Value.ToString());

                // iDate Total 6 mo
                dataGridView_iDate.Rows[3].Cells[2].Value = double.Parse(dataGridView_iDate.Rows[3].Cells[0].Value.ToString()) + double.Parse(dataGridView_iDate.Rows[3].Cells[1].Value.ToString());


                // iDate Total 1yr
                dataGridView_iDate.Rows[4].Cells[2].Value = double.Parse(dataGridView_iDate.Rows[4].Cells[0].Value.ToString()) + double.Parse(dataGridView_iDate.Rows[4].Cells[1].Value.ToString());


                //
                // ================================================================================================================================
                //

                // Date2 Firm pd
                DateTime[] keyDates_picker2 = { endDateTimePicker.Value, endDateTimePicker.Value.AddDays(30), endDateTimePicker.Value.AddDays(90), endDateTimePicker.Value.AddMonths(6), endDateTimePicker.Value.AddYears(1) };
                selectQuery = string.Empty;
                for (int i = 0; i < keyDates_picker2.Length; i++)
                    for (int j = 0; j < statuses.Length; j++)
                    {
                        selectQuery += "SUM(CASE WHEN (Status IN " + statuses[j] + " AND CONVERT(DATE, Promised_Date) <= CONVERT(DATE,'" + keyDates_picker2[i].ToShortDateString() + "')) THEN Order_Qty ELSE 0 END)";
                        if (j != statuses.Length - 1 || i != keyDates_picker2.Length - 1)
                            selectQuery += ", ";
                        else
                            selectQuery += " ";
                    }

                query =
                    "SELECT " + selectQuery +
                                    "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT " +
                                    "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) = " +
                                        "(SELECT CONVERT(DATETIME, MAX(InnerT.idate)) " +
                                        "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT " +
                                        "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + endDateTimePicker.Value.ToShortDateString() + "'));";

                com = new OdbcCommand(query, con);
                read = com.ExecuteReader();
                read.Read();
                for (int i = 0; i < keyDates_picker2.Length; i++)
                    for (int j = 0; j < statuses.Length; j++)
                    {
                        double orderQty = read.IsDBNull(i * statuses.Length + j) ? 0 : read.GetDouble(i * statuses.Length + j);
                        dataGridView_current.Rows[i].Cells[j].Value = orderQty;
                        if (orderQty > 0 && i == 0)
                            dataGridView_current.Rows[i].Cells[j].Style = style;
                        else
                            dataGridView_current.Rows[i].Cells[j].Style = dataGridView_current.Columns[j].DefaultCellStyle;
                    }

                // current Total pd
                total = double.Parse(dataGridView_current.Rows[0].Cells[0].Value.ToString()) + double.Parse(dataGridView_current.Rows[0].Cells[1].Value.ToString());
                dataGridView_current.Rows[0].Cells[2].Value = total;
                if (total > 0)
                    dataGridView_current.Rows[0].Cells[2].Style = style;
                else
                    dataGridView_current.Rows[0].Cells[2].Style = dataGridView_current.Columns[2].DefaultCellStyle;

                // current Total 30d
                dataGridView_current.Rows[1].Cells[2].Value = double.Parse(dataGridView_current.Rows[1].Cells[0].Value.ToString()) + double.Parse(dataGridView_current.Rows[1].Cells[1].Value.ToString());

                // current Total 90d
                dataGridView_current.Rows[2].Cells[2].Value = double.Parse(dataGridView_current.Rows[2].Cells[0].Value.ToString()) + double.Parse(dataGridView_current.Rows[2].Cells[1].Value.ToString());

                // current Total 6 mo
                dataGridView_current.Rows[3].Cells[2].Value = double.Parse(dataGridView_current.Rows[3].Cells[0].Value.ToString()) + double.Parse(dataGridView_current.Rows[3].Cells[1].Value.ToString());


                // current Total 1yr
                dataGridView_current.Rows[4].Cells[2].Value = double.Parse(dataGridView_current.Rows[4].Cells[0].Value.ToString()) + double.Parse(dataGridView_current.Rows[4].Cells[1].Value.ToString());
            }
        }

        private void Update_Monthly(object sender, EventArgs e)
        {
            if (iDate_yearDropDown.SelectedItem == null || linesDataGridView.Rows.Count == 0)
                return;

            linesDataGridView.Enabled = false;
            startDateTimePicker.Enabled = false;
            iDate_yearDropDown.Enabled = false;
            changesRadioButton.Enabled = false;
            openLinesRadioButton.Enabled = false;
            partsRadioButton.Enabled = false;
            customersRadioButton.Enabled = false;
            customerAndPartsRadioButton.Enabled = false;
            ByAllRadioButton.Enabled = false;
            addButton.Enabled = false;
            firmCheckBox.Enabled = false;
            shippedCheckBox.Enabled = false;
            forecastCheckBox.Enabled = false;
            selectAllCheckBox.Enabled = false;

            string part = linesDataGridView.SelectedRows[0].Cells[3].Value.ToString();
            int year = int.Parse(iDate_yearDropDown.SelectedItem.ToString());

            string[] statuses = { "('Open')", "('Hold')", "('Shipped')" };
            Boolean[] moneyFilters = { firmCheckBox.Checked, forecastCheckBox.Checked, shippedCheckBox.Checked };

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
            // IDATE
            //
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
                        selectString_Forecast += "SUM(CASE WHEN (Status IN " + statuses[j] + " AND CONVERT(DATE," + dateRangeVar + ") <= CONVERT(DATE, '" + (new DateTime(year, i, DateTime.DaysInMonth(year, i))).ToShortDateString() + "') AND CONVERT(DATE, " + dateRangeVar + ") >= CONVERT(DATE, '" + (new DateTime(year, i, 1)).ToShortDateString() + "')) THEN Order_Qty ELSE 0 END),\n" +
                                    "SUM(CASE WHEN (Status IN " + statuses[j] + " AND CONVERT(DATE, " + dateRangeVar + ") <= CONVERT(DATE, '" + (new DateTime(year, i, DateTime.DaysInMonth(year, i))).ToShortDateString() + "') AND CONVERT(DATE, " + dateRangeVar + ") >= CONVERT(DATE, '" + (new DateTime(year, i, 1)).ToShortDateString() + "')) THEN Total_Price ELSE 0 END)";
                        if (!(j == 1 && i == 12))
                            selectString_Forecast += ",\n";
                        else
                            selectString_Forecast += "\n";
                    }

                string query =
                    "SELECT " + selectString_Forecast + " " +
                            "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT\n" +
                            "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) =\n" +
                                "(SELECT CONVERT(DATETIME, MAX(InnerT.idate))\n" +
                                "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT\n" +
                                "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + startDateTimePicker.Value.ToShortDateString() + "'));\n";
                OdbcCommand com = new OdbcCommand(query, con);
                OdbcDataReader read = com.ExecuteReader();

                read.Read();
                for (int j = 0; j < 2; j++)
                    for (int i = 0; i < 12; i++)
                    {
                        double orderQty = read.IsDBNull((j * 12 + i) * 2) ? 0 : read.GetDouble((j * 12 + i) * 2);
                        quarter_iDateGridView.Rows[i].Cells[j].Value = orderQty;
                        totalCosts[i] += !read.IsDBNull((j * 12 + i) * 2 + 1) && moneyFilters[j] ? read.GetDecimal((j * 12 + i) * 2 + 1) : 0;
                    }
                read.Read();

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
                                "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND InnerT.Status = 'Shipped' AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + startDateTimePicker.Value.ToShortDateString() + "'));";
                com = new OdbcCommand(query, con);
                read = com.ExecuteReader();

                read.Read();
                for (int j = 2; j < 3; j++)
                    for (int i = 0; i < 12; i++)
                    {
                        double orderQty = read.IsDBNull(i * 2) ? 0 : read.GetDouble(i * 2);
                        quarter_iDateGridView.Rows[i].Cells[j].Value = orderQty;
                        totalCosts[i] += !read.IsDBNull(i * 2 + 1) && moneyFilters[j] ? read.GetDecimal(i * 2 + 1) : 0;
                    }

                read.Read();

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
                            "\tWHERE outT.SO_Detail = inTable.SO_Detail AND CONVERT(DATE, idate) <= CONVERT(DATE, '" + startDateTimePicker.Value.ToShortDateString() + "')); ";
                com = new OdbcCommand(query, con);
                read = com.ExecuteReader();

                read.Read();
                for (int i = 0; i < 12; i++)
                {
                    quarter_iDateGridView.Rows[i].Cells[4].Value = totalCosts[i];
                    quarter_iDateGridView.Rows[i].Cells[3].Value = read.IsDBNull(i) ? 0 : read.GetDouble(i);
                }
                read.Read();
                quarter_iDateGridView.Columns[4].DefaultCellStyle.Format = "c";

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
                                                "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + endDateTimePicker.Value.ToShortDateString() + "'));";


                com = new OdbcCommand(query, con);
                read = com.ExecuteReader();
                read.Read();
                for (int j = 0; j < 2; j++)
                    for (int i = 0; i < 12; i++)
                    {
                        double orderQty = read.IsDBNull((j * 12 + i) * 2) ? 0 : read.GetDouble((j * 12 + i) * 2);
                        quarter_currentGridView.Rows[i].Cells[j].Value = orderQty;
                        totalCosts[i] += !read.IsDBNull((j * 12 + i) * 2 + 1) && moneyFilters[j] ? read.GetDecimal((j * 12 + i) * 2 + 1) : 0;
                    }
                read.Read();


                // =========================================================

                query =
                    "SELECT " + selectString_Shipping + " " +
                            "FROM ATIDelivery.dbo.Delivery_Lines AS OutterT " +
                            "WHERE SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND " + filter + " AND CONVERT(DATETIME, idate) = " +
                                                "(SELECT CONVERT(DATETIME, MAX(InnerT.idate)) " +
                                                "FROM ATIDelivery.dbo.Delivery_Lines AS InnerT " +
                                                "WHERE InnerT.SO_Detail = OutterT.SO_Detail AND InnerT.Status = 'Shipped' AND CONVERT(DATE,InnerT.idate) <= CONVERT(DATE, '" + endDateTimePicker.Value.ToShortDateString() + "'));";


                com = new OdbcCommand(query, con);
                read = com.ExecuteReader();
                read.Read();
                for (int j = 2; j < 3; j++)
                    for (int i = 0; i < 12; i++)
                    {
                        double orderQty = read.IsDBNull(i * 2) ? 0 : read.GetDouble(i * 2);
                        quarter_currentGridView.Rows[i].Cells[j].Value = orderQty;
                        totalCosts[i] += !read.IsDBNull(i * 2 + 1) && moneyFilters[j] ? read.GetDecimal(i * 2 + 1) : 0;
                    }
                read.Read();

                //==============================================================================

                query = "SELECT " + selectString_totalDue + " " +
                        "FROM ATIDelivery.dbo.Delivery_Lines AS outT\n" +
                        "WHERE " + filter + " AND Status <> 'Removed' AND SO_Line NOT IN ('LAIR', 'FAIR', 'NREC', 'BTS') AND  CONVERT(DATETIME, idate) =\n" +
                            "\t(SELECT CONVERT(DATETIME, MAX(inTable.idate))\n" +
                            "\tFROM ATIDelivery.dbo.Delivery_Lines AS inTable\n" +
                            "\tWHERE outT.SO_Detail = inTable.SO_Detail AND CONVERT(DATETIME, idate) <= CONVERT(DATE, '" + endDateTimePicker.Value.ToShortDateString() + "')); ";
                com = new OdbcCommand(query, con);
                read = com.ExecuteReader();

                read.Read();
                for (int i = 0; i < 12; i++)
                {
                    quarter_currentGridView.Rows[i].Cells[4].Value = totalCosts[i];
                    quarter_currentGridView.Rows[i].Cells[3].Value = read.IsDBNull(i) ? 0 : read.GetDouble(i); ;
                }
                quarter_currentGridView.Columns[4].DefaultCellStyle.Format = "c";
                read.Read();
            }
            // ==========================================================


            linesDataGridView.Enabled = true;
            startDateTimePicker.Enabled = true;
            iDate_yearDropDown.Enabled = true;
            changesRadioButton.Enabled = true;
            openLinesRadioButton.Enabled = true;
            partsRadioButton.Enabled = true;
            customersRadioButton.Enabled = true;
            customerAndPartsRadioButton.Enabled = true;
            ByAllRadioButton.Enabled = true;
            addButton.Enabled = true;
            firmCheckBox.Enabled = true;
            shippedCheckBox.Enabled = true;
            forecastCheckBox.Enabled = true;
            selectAllCheckBox.Enabled = true;

        }

        private void selectAllCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.firmCheckBox.CheckedChanged -= this.firmCheckBox_CheckedChanged;
            this.firmCheckBox.CheckedChanged -= this.Update_Monthly;
            this.forecastCheckBox.CheckedChanged -= this.forecastCheckBox_CheckedChanged;
            this.forecastCheckBox.CheckedChanged -= this.Update_Monthly;
            this.shippedCheckBox.CheckedChanged -= this.totalCheckBox_CheckedChanged;
            this.shippedCheckBox.CheckedChanged -= this.Update_Monthly;
            firmCheckBox.Checked = true;
            forecastCheckBox.Checked = true;
            shippedCheckBox.Checked = true;
            this.firmCheckBox.CheckedChanged += this.firmCheckBox_CheckedChanged;
            this.firmCheckBox.CheckedChanged += this.Update_Monthly;
            this.forecastCheckBox.CheckedChanged += this.forecastCheckBox_CheckedChanged;
            this.forecastCheckBox.CheckedChanged += this.Update_Monthly;
            this.shippedCheckBox.CheckedChanged += this.totalCheckBox_CheckedChanged;
            this.shippedCheckBox.CheckedChanged += this.Update_Monthly;
        }

        private void firmCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.selectAllCheckBox.CheckedChanged -= this.selectAllCheckBox_CheckedChanged;
            this.selectAllCheckBox.CheckedChanged -= this.Update_Monthly;
            selectAllCheckBox.Checked = false;
            this.selectAllCheckBox.CheckedChanged += this.selectAllCheckBox_CheckedChanged;
            this.selectAllCheckBox.CheckedChanged += this.Update_Monthly;
        }

        private void forecastCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.selectAllCheckBox.CheckedChanged -= this.selectAllCheckBox_CheckedChanged;
            this.selectAllCheckBox.CheckedChanged -= this.Update_Monthly;
            selectAllCheckBox.Checked = false;
            this.selectAllCheckBox.CheckedChanged += this.selectAllCheckBox_CheckedChanged;
            this.selectAllCheckBox.CheckedChanged += this.Update_Monthly;
        }

        private void totalCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            this.selectAllCheckBox.CheckedChanged -= this.selectAllCheckBox_CheckedChanged;
            this.selectAllCheckBox.CheckedChanged -= this.Update_Monthly;
            selectAllCheckBox.Checked = false;
            this.selectAllCheckBox.CheckedChanged += this.selectAllCheckBox_CheckedChanged;
            this.selectAllCheckBox.CheckedChanged += this.Update_Monthly;
        }

        private void dgvSomeDataGridView_SelectionChanged(Object sender, EventArgs e)
        {
            ((DataGridView)sender).ClearSelection();
        }

        private void Form1_HelpButtonClicked(object sender, CancelEventArgs e)
        {
            System.Diagnostics.Process.Start(@"T:\users\Daan Leiva\ProjectsBackUp\so_iNTERFACE\segd\SO_Delivery_Interface\SO_Delivery_Interface\bin\Release\helpHolder.pdf");
        }
    }
}

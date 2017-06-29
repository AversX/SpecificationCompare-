using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Drawing;
using System.Text;
using System.Text.RegularExpressions;

namespace SpecificationCompare
{
    struct Unit
    {
        List<string> columns;
        List<Error> errors;
        bool finded;

        public List<string> Columns { get => columns; set => columns = value; }
        public bool Finded { get => finded; set => finded = value; }
        internal List<Error> Errors { get => errors; set => errors = value; }
    }

    struct Error
    {
        string errorCode;
        string oldValue;
        int colNumber;

        public string ErrorCode { get => errorCode; set => errorCode = value; }
        public string OldValue { get => oldValue; set => oldValue = value; }
        public int ColNumber { get => colNumber; set => colNumber = value; }
    }

    public partial class MainForm : MetroFramework.Forms.MetroForm
    {
        private Excel.Application excel;
        private int artCol, nameCol, colNum;


        public MainForm()
        {
            InitializeComponent();
        }

        private void OldVerTxt_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && ((e.AllowedEffect & DragDropEffects.Move) == DragDropEffects.Move))
                e.Effect = DragDropEffects.Move;
        }

        private void OldVerTxt_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && e.Effect == DragDropEffects.Move)
            {
                string objects = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
                oldVerTxt.Text = objects;
            }
        }

        private void NewVerTxt_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop) && ((e.AllowedEffect & DragDropEffects.Move) == DragDropEffects.Move))
                e.Effect = DragDropEffects.Move;
        }

        private void NewVerTxt_DragDrop(object sender, DragEventArgs e)
        {
            string objects = ((string[])e.Data.GetData(DataFormats.FileDrop))[0];
            newVerTxt.Text = objects;
        }

        private void SrchSpecOldBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Filter = "(*.xlsx); (*.xls)|*.xlsx; *.xls";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                oldVerTxt.Text = ofd.FileName;
        }

        private void SrchSpecNewBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Multiselect = true;
            ofd.Filter = "(*.xlsx); (*.xls)|*.xlsx; *.xls";
            if (ofd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                newVerTxt.Text = ofd.FileName;
        }

        private void CompareBtn_Click(object sender, EventArgs e)
        {
           
            if (artColTBox.Text!="" && nameColTBox.Text != "")
            {
                if (artColTBox.Text != "") artCol = int.Parse(artColTBox.Text) - 1;
                if (nameColTBox.Text != "") nameCol = int.Parse(nameColTBox.Text) - 1;

                List<Unit> Units = new List<Unit>();
                
                List<Unit> OldUnits = new List<Unit>();
                OldUnits.AddRange(LoadDataSpec(oldVerTxt.Text));

                List<Unit> NewUnits = new List<Unit>();
                NewUnits.AddRange(LoadDataSpec(newVerTxt.Text));

                metroProgressBar1.Visible = true;
                metroProgressBar1.Minimum = 0;
                metroProgressBar1.Value = 0;

                Units = Consolidate(OldUnits, NewUnits);
                metroProgressBar1.Maximum = Units.Count;
                UploadData(Units);
            }
        }

        private List<Unit> LoadDataSpec(string path)
        {
            List<Unit> units = new List<Unit>();
            DataSet dataSet = new DataSet("EXCEL");
            string connectionString;
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties='Excel 12.0;" + (headerChBox.Checked ? "HDR=YES;IMEX=0;" : "HDR=NO;IMEX=1;") + "'";
            OleDbConnection connection = new OleDbConnection(connectionString);
            connection.Open();

            System.Data.DataTable schemaTable = connection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
            string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];

            string select = String.Format("SELECT * FROM [{0}]", sheet1);
            OleDbDataAdapter adapter = new OleDbDataAdapter(select, connection);
            adapter.Fill(dataSet);
            connection.Close();

            for (int row = 0; row < dataSet.Tables[0].Rows.Count; row++)
            {
                if (dataSet.Tables[0].Rows[row][artCol].ToString().Length > 0 || dataSet.Tables[0].Rows[row][nameCol].ToString().Length > 0)
                {
                    Unit unit = new Unit();
                    unit.Columns = new List<string>();
                    unit.Errors = new List<Error>();
                    unit.Finded = false;
                    for (int j = 0; j < dataSet.Tables[0].Columns.Count; j++)
                        unit.Columns.Add(dataSet.Tables[0].Rows[row][j].ToString().Trim());
                    while (unit.Columns[unit.Columns.Count-1] == "") unit.Columns.RemoveAt(unit.Columns.Count - 1);
                    units.Add(unit);
                }
            }
            colNum = units[0].Columns.Count;
            return units;
        }

        private List<Unit> Consolidate(List<Unit> oldUnits, List<Unit> newUnits)
        {
            for (int i = 0; i < oldUnits.Count; i++)
                for (int j = 0; j < newUnits.Count; j++)
                {
                    if (oldUnits[i].Columns[artCol] != "")
                    {
                        if (oldUnits[i].Columns[artCol] == newUnits[j].Columns[artCol])
                        {
                            if (oldUnits[i].Finded)
                            {
                                oldUnits.Insert(i, newUnits[j]);
                                Unit unit = oldUnits[i];
                                List<Error> errors = unit.Errors;
                                errors.Add(new Error() { ErrorCode = "REPEAT" });
                                if (unit.Errors == null) unit.Errors = new List<Error>(errors);
                                else unit.Errors = errors;
                                oldUnits[i] = unit;
                                i++;
                                newUnits.RemoveAt(j);
                                j--;
                            }
                            else
                            {
                                Unit unit = oldUnits[i];
                                List<Error> errors = oldUnits[i].Errors;
                                for (int c = 0; c < oldUnits[i].Columns.Count; c++)
                                {
                                    if (oldUnits[i].Columns[c].Replace(".", "") != newUnits[j].Columns[c].Replace(".", ""))
                                    {
                                        errors.Add(new Error
                                        {
                                            ErrorCode = "CHANGE",
                                            ColNumber = c,
                                            OldValue = oldUnits[i].Columns[c]
                                        });
                                        unit.Columns[c] = newUnits[j].Columns[c];
                                    }
                                }
                                unit.Errors = errors;
                                unit.Finded = true;
                                oldUnits[i] = unit;
                                newUnits.RemoveAt(j);
                                j--;
                            }
                        }
                        else if (oldUnits[i].Columns[nameCol] == newUnits[j].Columns[nameCol])
                        {
                            if (oldUnits[i].Finded)
                            {
                                oldUnits.Insert(i, newUnits[j]);
                                Unit unit = oldUnits[i];
                                List<Error> errors = new List<Error>();
                                errors.Add(new Error() { ErrorCode = "REPEAT" });
                                if (unit.Errors == null) unit.Errors = new List<Error>(errors);
                                else unit.Errors.AddRange(errors);
                                oldUnits[i] = unit;
                                i++;
                                newUnits.RemoveAt(j);
                                j--;
                            }
                            else
                            {
                                Unit unit = oldUnits[i];
                                List<Error> errors = oldUnits[i].Errors;
                                for (int c = 0; c < oldUnits[i].Columns.Count; c++)
                                {
                                    if (oldUnits[i].Columns[c].Replace(".", "") != newUnits[j].Columns[c].Replace(".", ""))
                                    {
                                        errors.Add(new Error
                                        {
                                            ErrorCode = "CHANGE",
                                            ColNumber = c,
                                            OldValue = oldUnits[i].Columns[c]
                                        });
                                        unit.Columns[c] = newUnits[j].Columns[c];
                                    }
                                }
                                unit.Errors = errors;
                                unit.Finded = true;
                                oldUnits[i] = unit;
                                newUnits.RemoveAt(j);
                                j--;
                            }
                        }
                    }
                    else
                    {
                        if (oldUnits[i].Columns[nameCol] == newUnits[j].Columns[nameCol])
                        {
                            if (oldUnits[i].Finded)
                            {
                                oldUnits.Insert(i, newUnits[j]);
                                Unit unit = oldUnits[i];
                                List<Error> errors = unit.Errors;
                                errors.Add(new Error() { ErrorCode = "REPEAT" });
                                if (unit.Errors == null) unit.Errors = new List<Error>(errors);
                                else unit.Errors = errors;
                                oldUnits[i] = unit;
                                newUnits.RemoveAt(j);
                                j--;
                            }
                            else
                            {
                                Unit unit = oldUnits[i];
                                List<Error> errors = oldUnits[i].Errors;
                                for (int c = 0; c < oldUnits[i].Columns.Count; c++)
                                {
                                    if (oldUnits[i].Columns[c].Replace(".", "") != newUnits[j].Columns[c].Replace(".", ""))
                                    {
                                        errors.Add(new Error
                                        {
                                            ErrorCode = "CHANGE",
                                            ColNumber = c,
                                            OldValue = oldUnits[i].Columns[c]
                                        });
                                        unit.Columns[c] = newUnits[j].Columns[c];
                                    }
                                }
                                unit.Errors = errors;
                                unit.Finded = true;
                                oldUnits[i] = unit;
                                newUnits.RemoveAt(j);
                                j--;
                            }
                        }
                    }
                }
            
            for (int j = 0; j < newUnits.Count; j++)
            {
                Unit unit = newUnits[j];
                List<Error> errors = unit.Errors;
                errors.Add(new Error() { ErrorCode = "NEW" });
                if (unit.Errors == null) unit.Errors = new List<Error>(errors);
                else unit.Errors.AddRange(errors);
                oldUnits.Add(unit);
            }
            return oldUnits;
        }

        private void UploadData(List<Unit> units)
        {
            excel = new Excel.Application();
            excel.SheetsInNewWorkbook = 1;
            excel.Workbooks.Add(Type.Missing);
            Excel.Worksheet sheet = (Excel.Worksheet)excel.Sheets.get_Item(1);
            Excel.Range autoFit;

            int curColumn = 1;
            for (int i = 1; i <= colNum; i++)
            {
                sheet.Columns[curColumn].NumberFormat = "@";
                curColumn++;
            }

            for (int i = 0; i < units.Count; i++)
            {
                for (int j = 0; j < colNum; j++)
                {
                    sheet.Cells[i + 1, j+1] = units[i].Columns[j];
                }
                autoFit = (Excel.Range)sheet.Rows[i + 1];
                autoFit.EntireRow.AutoFit();
                for (int j = 1; j <= curColumn - 1; j++)
                {
                    autoFit = (Excel.Range)sheet.Columns[j];
                    autoFit.AutoFit();
                }
                if (units[i].Errors != null)
                {
                    for (int j = 0; j < units[i].Errors.Count; j++)
                    {
                        switch (units[i].Errors[j].ErrorCode)
                        {
                            case "REPEAT": { ((Excel.Range)sheet.Rows[i + 1]).EntireRow.Interior.Color = Color.Blue; break; }
                            case "NEW": { ((Excel.Range)sheet.Rows[i + 1]).EntireRow.Interior.Color = Color.Green; break; }
                            default:
                                {
                                    if (!buyChBox.Checked)
                                    {
                                        sheet.Cells[i + 1, units[i].Errors[j].ColNumber + 1].Interior.Color = Color.Red;
                                        ((Excel.Range)sheet.Cells[i + 1, units[i].Errors[j].ColNumber + 1]).AddComment(units[i].Errors[j].OldValue);
                                    }
                                    break;
                                }
                        }
                    }
                }
                else
                {
                    ((Excel.Range)sheet.Rows[i + 1]).EntireRow.Interior.Color = Color.Yellow;
                }
                metroProgressBar1.Value++;
            }
            excel.Visible = true;
            metroProgressBar1.Visible = false;
        }

        private void columnsTBox_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;
            if (!Char.IsDigit(number) && number != 8)// && number != 44)
            {
                e.Handled = true;
            }
        }
    }
}

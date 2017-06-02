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
<<<<<<< HEAD
        List<string> columns;
=======
        string pos;
        string article;
        string name;
        string number;
        string measure;
        string manufacture;


>>>>>>> 2f4e78a3154cfa65e2d46d0334ab19cfbd148bf8
        List<Error> errors;
        bool finded;

<<<<<<< HEAD
        public List<string> Columns { get => columns; set => columns = value; }
=======
        public string Manufacture { get => manufacture; set => manufacture = value; }
        public string Measure { get => measure; set => measure = value; }
        public string Name { get => name; set => name = value; }
        public string Article { get => article; set => article = value; }
        public string Pos { get => pos; set => pos = value; }
        public string Number { get => number; set => number = value; }
>>>>>>> 2f4e78a3154cfa65e2d46d0334ab19cfbd148bf8
        public bool Finded { get => finded; set => finded = value; }
        internal List<Error> Errors { get => errors; set => errors = value; }
    }

    struct Error
    {
        int errorCode;
        string oldValue;

        public int ErrorCode { get => errorCode; set => errorCode = value; }
        public string OldValue { get => oldValue; set => oldValue = value; }
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

                Units = Consolidate(OldUnits, NewUnits);
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

            colNum = dataSet.Tables[0].Columns.Count;
            for (int row = 0; row < dataSet.Tables[0].Rows.Count; row++)
            {
                if (dataSet.Tables[0].Rows[row][artCol].ToString().Length > 0 || dataSet.Tables[0].Rows[row][nameCol].ToString().Length > 0)
                {
<<<<<<< HEAD
                    Unit unit = new Unit();
                    unit.Columns = new List<string>();
                    unit.Errors = new List<Error>();
                    unit.Finded = false;
                    for (int j = 0; j < dataSet.Tables[0].Columns.Count; j++)
                        unit.Columns.Add(dataSet.Tables[0].Rows[row][j].ToString().Trim());
=======
                    Unit unit = new Unit()
                    {
                        Pos = dataSet.Tables[0].Rows[row][1].ToString().Trim(),
                        Article = dataSet.Tables[0].Rows[row][2].ToString().Trim(),
                        Name = dataSet.Tables[0].Rows[row][3].ToString().Trim(),
                        Number = dataSet.Tables[0].Rows[row][4].ToString().Trim(),
                        Measure = dataSet.Tables[0].Rows[row][5].ToString().Trim(),
                        Manufacture = dataSet.Tables[0].Rows[row][6].ToString().Trim(),
                        Finded = false
                    };
>>>>>>> 2f4e78a3154cfa65e2d46d0334ab19cfbd148bf8
                    units.Add(unit);
                }
            }
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
                                errors.Add(new Error() { ErrorCode = 6 });
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
<<<<<<< HEAD
                                List<Error> errors = oldUnits[i].Errors;
                                for (int c = 0; c < oldUnits[i].Columns.Count; c++)
=======
                                List<Error> errors = new List<Error>();
                                if (oldUnits[i].Pos != newUnits[j].Pos)
                                {
                                    errors.Add(new Error
                                    {
                                        ErrorCode = 0,
                                        OldValue = oldUnits[i].Pos
                                    });
                                    unit.Pos = newUnits[j].Pos;
                                }
                                if (oldUnits[i].Name != newUnits[j].Name)
                                {
                                    errors.Add(new Error
                                    {
                                        ErrorCode = 2,
                                        OldValue = oldUnits[i].Name
                                    });
                                    unit.Name = newUnits[j].Name;
                                }
                                if (oldUnits[i].Number != newUnits[j].Number)
>>>>>>> 2f4e78a3154cfa65e2d46d0334ab19cfbd148bf8
                                {
                                    if (oldUnits[i].Columns[c].Replace(".", "") != newUnits[j].Columns[c].Replace(".", ""))
                                    {
                                        errors.Add(new Error
                                        {
                                            ErrorCode = c,
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
                                errors.Add(new Error() { ErrorCode = 6 });
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
<<<<<<< HEAD
                                List<Error> errors = oldUnits[i].Errors;
                                for (int c = 0; c < oldUnits[i].Columns.Count; c++)
=======
                                List<Error> errors = new List<Error>();
                                if (oldUnits[i].Pos != newUnits[j].Pos)
>>>>>>> 2f4e78a3154cfa65e2d46d0334ab19cfbd148bf8
                                {
                                    if (oldUnits[i].Columns[c].Replace(".", "") != newUnits[j].Columns[c].Replace(".", ""))
                                    {
<<<<<<< HEAD
                                        errors.Add(new Error
                                        {
                                            ErrorCode = c,
                                            OldValue = oldUnits[i].Columns[c]
                                        });
                                        unit.Columns[c] = newUnits[j].Columns[c];
                                    }
=======
                                        ErrorCode = 0,
                                        OldValue = oldUnits[i].Pos
                                    });
                                    unit.Pos = newUnits[j].Pos;
>>>>>>> 2f4e78a3154cfa65e2d46d0334ab19cfbd148bf8
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
                                errors.Add(new Error() { ErrorCode = 6 });
                                if (unit.Errors == null) unit.Errors = new List<Error>(errors);
                                else unit.Errors = errors;
                                oldUnits[i] = unit;
                                newUnits.RemoveAt(j);
                                j--;
                            }
                            else
                            {
                                Unit unit = oldUnits[i];
<<<<<<< HEAD
                                List<Error> errors = oldUnits[i].Errors;
                                for (int c = 0; c < oldUnits[i].Columns.Count; c++)
=======
                                List<Error> errors = new List<Error>();
                                if (oldUnits[i].Pos != newUnits[j].Pos)
                                {
                                    errors.Add(new Error
                                    {
                                        ErrorCode = 0,
                                        OldValue = oldUnits[i].Pos
                                    });
                                    unit.Pos = newUnits[j].Pos;
                                }
                                if (oldUnits[i].Article != newUnits[j].Article)
                                {
                                    errors.Add(new Error
                                    {
                                        ErrorCode = 1,
                                        OldValue = oldUnits[i].Article
                                    });
                                    unit.Name = newUnits[j].Name;
                                }
                                if (oldUnits[i].Number != newUnits[j].Number)
                                {
                                    errors.Add(new Error
                                    {
                                        ErrorCode = 3,
                                        OldValue = oldUnits[i].Number
                                    });
                                    unit.Number = newUnits[j].Number;
                                }
                                if (oldUnits[i].Measure.Replace(".", "") != newUnits[j].Measure.Replace(".", ""))
                                {
                                    errors.Add(new Error
                                    {
                                        ErrorCode = 4,
                                        OldValue = oldUnits[i].Measure
                                    });
                                    unit.Measure = newUnits[j].Measure;
                                }
                                if (oldUnits[i].Manufacture != newUnits[j].Manufacture)
>>>>>>> 2f4e78a3154cfa65e2d46d0334ab19cfbd148bf8
                                {
                                    if (oldUnits[i].Columns[c].Replace(".", "") != newUnits[j].Columns[c].Replace(".", ""))
                                    {
                                        errors.Add(new Error
                                        {
                                            ErrorCode = c,
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
                errors.Add(new Error() { ErrorCode = 7 });
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
<<<<<<< HEAD
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
=======

            //sheet.Cells[1, curColumn] = "Поз.";
            sheet.Columns[curColumn].NumberFormat = "@";
            curColumn++;

            //sheet.Cells[1, curColumn] = "Код";
            sheet.Columns[curColumn].NumberFormat = "@";
            curColumn++;

            //sheet.Cells[1, curColumn] = "Наименование";
            sheet.Columns[curColumn].NumberFormat = "@";
            curColumn++;

            //sheet.Cells[1, curColumn] = "Кол-во";
            sheet.Columns[curColumn].NumberFormat = "@";
            curColumn++;

            //sheet.Cells[1, curColumn] = "Ед. изм.";
            sheet.Columns[curColumn].NumberFormat = "@";
            curColumn++;

            //sheet.Cells[1, curColumn] = "Завод изготовитель";
            sheet.Columns[curColumn].NumberFormat = "@";
            curColumn++;

            for (int i = 0; i < units.Count; i++)
            {
                sheet.Cells[i + 1, curColumn - 6] = units[i].Pos;
                sheet.Cells[i + 1, curColumn - 5] = units[i].Article;
                sheet.Cells[i + 1, curColumn - 4] = units[i].Name;
                sheet.Cells[i + 1, curColumn - 3] = units[i].Number;
                sheet.Cells[i + 1, curColumn - 2] = units[i].Measure;
                sheet.Cells[i + 1, curColumn - 1] = units[i].Manufacture;
>>>>>>> 2f4e78a3154cfa65e2d46d0334ab19cfbd148bf8
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
                            case 6: { ((Excel.Range)sheet.Rows[i + 1]).EntireRow.Interior.Color = Color.Blue; break; }
                            case 7: { ((Excel.Range)sheet.Rows[i + 1]).EntireRow.Interior.Color = Color.Green; break; }
                            default:
                                {
                                    sheet.Cells[i + 1, units[i].Errors[j].ErrorCode + 1].Interior.Color = Color.Red;
                                    ((Excel.Range)sheet.Cells[i + 1, units[i].Errors[j].ErrorCode + 1]).AddComment(units[i].Errors[j].OldValue);
                                    break;
                                }
                        }
                    }
                }
                else
                {
                    ((Excel.Range)sheet.Rows[i + 1]).EntireRow.Interior.Color = Color.Yellow;
                }

            }
            excel.Visible = true;
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

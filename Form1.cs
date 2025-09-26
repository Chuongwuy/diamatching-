using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Media;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace diematching
{

    public partial class Diematching : Form
    {
        private string selectedFilePath = "";
        string outputExcel;
        Dictionary<(int row, int col), DataTable> popupDict = new Dictionary<(int row, int col), DataTable>();
        void UpdateStatus(string message)
        {
            // Ví dụ cập nhật một Label để hiển thị trạng thái
            resultlabel.Text = message;
        }


        public Diematching()
        {
            InitializeComponent();
            ExcelPackage.License.SetNonCommercialPersonal("HUY");    //License NET 4.8

            browsebutton.Click += (s, e) => BrowseFile();
            generatebutton.Click += (s, e) => GenerateFiles();
            cancelbutton.Click += (s, e) => Application.Exit();
            loadbutton.Click += (s, e) => LoadFiles();
            savebutton.Click += (s, e) => Browseoutput();
            datagridview.CellClick += CellClick;
        }


        //-----------------------------------------------------------
        private void button1_Click(object sender, EventArgs e)
        {

        }
        private void resultlabel_Click(object sender, EventArgs e)
        {

        }
        private void Diematching_Load(object sender, EventArgs e)
        {

        }
        //------------------------------------------------------------

        private void BrowseFile()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|All Files|*.*";
                openFileDialog.Title = "Select Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    selectedFilePath = openFileDialog.FileName;
                    rawtextbox.Text = selectedFilePath;             // write the filepath into textbox 
                    UpdateStatus("File selected: " + Path.GetFileName(selectedFilePath));
                }
            }
        }

        private void Browseoutput()
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|All Files|*.*";
                openFileDialog.Title = "Select Excel File";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    outputExcel = openFileDialog.FileName;
                    savetextbox.Text = outputExcel;             // write the filepath into textbox 
                    UpdateStatus("File selected: " + Path.GetFileName(outputExcel));
                }
            }
            //using (FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog())
            //{
            //    folderBrowserDialog.Description = "Select Output Folder";
            //    if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            //    {
            //        // Lưu đường dẫn thư mục đã chọn
            //        outputExcel = folderBrowserDialog.SelectedPath;
            //        savetextbox.Text = outputExcel;

            //        UpdateStatus("Folder selected: " + outputExcel);
            //    }
            //}
        }


        private void GenerateFiles()
        {
            if (string.IsNullOrEmpty(selectedFilePath))
            {
                MessageBox.Show($"Please select an input data first", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (string.IsNullOrEmpty(outputExcel))
            {
                MessageBox.Show($"Please select an output data", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            try
            {
                // Kiểm tra file tồn tại
                if (!File.Exists(selectedFilePath))
                {
                    MessageBox.Show($"File input not found at {selectedFilePath} ", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (!File.Exists(outputExcel))
                {
                    MessageBox.Show($"File output not found at {outputExcel} ", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                //--------------------------PROCESS----------------------------
                //string timestamp = DateTime.Now.ToString("HH-mm-ss__dd-MM-yyyy");
                //string outputFileName = $"diematching_{timestamp}.xlsx";
                //string output = Path.Combine(outputExcel, outputFileName);

                var fileInfodata = new FileInfo(selectedFilePath);
                var fileInfomatrix = new FileInfo(outputExcel);
                using (var packageData = new ExcelPackage(fileInfodata))                 // access to file raw data P1,P2,...C1,C2
                using (var packageMatrix = new ExcelPackage(fileInfomatrix))             // access to file main matrix, calculate, pass fail
                {

                    var worksheetPlug = packageData.Workbook.Worksheets["P"];
                    if (worksheetPlug == null)
                    {
                        Console.WriteLine("Can not find any worksheet named P in input data file. Please check the name of worksheet.");
                        return;
                    }

                    var worksheetCeramic = packageData.Workbook.Worksheets["C"];
                    if (worksheetCeramic == null)
                    {
                        Console.WriteLine("Can not find any worksheet named C in input data file. Please check the name of worksheet.");
                        return;
                    }



                    //-------------------------ADD WORKSHEET 'MAIN'----------------------------------

                    var worksheetMatrix = packageMatrix.Workbook.Worksheets["MAIN"];
                    if (worksheetMatrix == null)
                    {
                        worksheetMatrix = packageMatrix.Workbook.Worksheets.Add("MAIN");

                    }
                    worksheetMatrix.Cells.Clear();      // clear old data

                    // -----------------------------------------CHECK STOCK PLUG AND LOT PLUG ----------------------------------------


                    int plug = 0;
                    for (int stockP = 11; stockP <= worksheetPlug.Dimension.End.Row; stockP++)
                    {
                        var cellstockP = worksheetPlug.Cells[stockP, 15];
                        if (cellstockP?.Value != null && cellstockP.Value.ToString() != "0" && double.TryParse(cellstockP.Value.ToString(), out double Pvalue))  // all cases approved
                        {
                            for (int plugcheck = 11; plugcheck <= worksheetPlug.Dimension.End.Row; plugcheck++)
                            {
                                var cellplug = worksheetPlug.Cells[plugcheck, 3];
                                if (cellplug.Value != null && !string.IsNullOrEmpty(cellplug.Value.ToString()))
                                {
                                    plug++;
                                    break;
                                }
                            }
                        }

                    }
                    worksheetMatrix.Cells["A1"].Value = "Lot plug";
                    worksheetMatrix.Cells["A2"].Value = plug;
                    var colorplug = worksheetMatrix.Cells["A2"];
                    colorplug.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    colorplug.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);

                    // ----------------------------------CHECK STOCK CERAMIC AND LOT CERAMIC ----------------------------------------

                    int ceramic = 0;
                    for (int stockC = 11; stockC <= worksheetCeramic.Dimension.End.Row; stockC++)
                    {
                        var cellstockC = worksheetCeramic.Cells[stockC, 15];
                        if (cellstockC?.Value != null && !string.IsNullOrEmpty(cellstockC.Value.ToString()) && cellstockC.Value.ToString() != "0")
                        {
                            for (int ceramiccheck = 11; ceramiccheck <= worksheetCeramic.Dimension.End.Row; ceramiccheck++)
                            {
                                var cellceramic = worksheetCeramic.Cells[ceramiccheck, 3];
                                if (cellceramic.Value != null && !string.IsNullOrEmpty(cellceramic.Value.ToString()))
                                {
                                    ceramic++;
                                    break;
                                }
                            }
                        }
                    }
                    worksheetMatrix.Cells["B1"].Value = "Lot ceramic";
                    worksheetMatrix.Cells["B2"].Value = ceramic;
                    var colorceramic = worksheetMatrix.Cells["B2"];
                    colorceramic.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    colorceramic.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);


                    //----------------------------DELETE OLD UNRELEVANT WORKSHEET --------------------------


                    // Get all worksheet names except MAIN
                    var clearOLD = packageMatrix.Workbook.Worksheets
                        .Select(ws => ws.Name)
                        .Where(name => !string.Equals(name, "MAIN", StringComparison.OrdinalIgnoreCase))
                        .ToList();

                    // Delete by name
                    foreach (var sheetName in clearOLD)
                    {
                        packageMatrix.Workbook.Worksheets.Delete(sheetName);
                    }

                    // -----------------------------ADD WORKSHEET MATRIX -------------------------------------
                    // Fast look Up checking
                    var existingSheets = packageMatrix.Workbook.Worksheets
                    .Select(ws => ws.Name)
                    .ToHashSet(); // for fast lookup

                    // check name at array P and C
                    for (int plugcheck = 11; plugcheck <= worksheetPlug.Dimension.End.Row; plugcheck++)
                    {
                        for (int ceramiccheck = 11; ceramiccheck <= worksheetCeramic.Dimension.End.Row; ceramiccheck++)
                        {
                            var cellC = worksheetCeramic.Cells[ceramiccheck, 3].Value;
                            var cellP = worksheetPlug.Cells[plugcheck, 3].Value;

                            string sheetName = $"{cellP}_{cellC}";
                            // Only add if NOT already exists
                            if (!existingSheets.Contains(sheetName))
                            {
                                packageMatrix.Workbook.Worksheets.Add(sheetName);
                            }

                        }
                    }
                    packageMatrix.SaveAs(new FileInfo(outputExcel));



                    // ---------------------COPY PLUG -------------------------------

                    int copyplug = 11;
                    for (int stockP = 11; stockP <= worksheetPlug.Dimension.End.Row; stockP++)
                    {
                        var cellstockP = worksheetPlug.Cells[stockP, 15];

                        if (cellstockP?.Value == null || !double.TryParse(cellstockP.Value.ToString(), out double stockvalue) || stockvalue == 0)  // all cases invalid
                        {
                            //Console.WriteLine($"cell null at ({stockP}, 15)  skip plug lot {worksheetPlug.Cells[stockP, 3].Value}");


                            // delete that worksheet if the lot plug name skipping
                            string deletename = worksheetPlug.Cells[stockP, 3].Value.ToString();

                            var deleteWS = packageMatrix.Workbook.Worksheets
                             .Where(ws => ws.Name.StartsWith(deletename + "_"))
                             .ToList();
                            foreach (var deletesheet in deleteWS)
                            {
                                packageMatrix.Workbook.Worksheets.Delete(deletesheet);
                                //Console.WriteLine($"plug lot worksheet {worksheetPlug.Cells[stockP, 3].Value} deleted");
                            }
                            copyplug++;
                            continue;
                        }

                        //// only process if stock # 0, empty
                        // stock # 0
                        if (copyplug <= worksheetPlug.Dimension.End.Row)
                        {

                            var sourceplug = worksheetPlug.Cells[copyplug, 3].Value;
                            if (sourceplug == null)
                            {
                                //Console.WriteLine($"Source plug is null at copyplug={copyplug}. Skipping.");
                                continue;
                            }
                            string prefixPLUG = sourceplug.ToString();   // got prefix sourceplug

                            // find target sheets to use that same prefix 
                            var ALLPLUG = packageMatrix.Workbook.Worksheets;
                            var targets = ALLPLUG
                                .Where(ws => ws.Name.StartsWith(prefixPLUG + "_"))
                                .ToList();
                            //Console.WriteLine($"Prefix plug is : {prefixPLUG}");

                            foreach (var targetsheet in targets)    // in target sheet need to copy 
                            {
                                //Console.WriteLine($" Found plug sheet: {targetsheet.Name}");

                                // range data need to copy
                                for (int i = 0; i < 11; i++)
                                {
                                    var sourcevalue = worksheetPlug.Cells[copyplug, 4 + i].Value;   // from 1 source worksheetPlug 
                                    var targetCell = targetsheet.Cells[11 + i, 3];
                                    targetCell.Value = sourcevalue;
                                    targetCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    targetCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                                }
                            }

                        }
                        copyplug++;
                    }







                    // ---------------------COPY CERAMIC -------------------------------

                    int copyceramic = 11;
                    for (int stockC = 11; stockC <= worksheetCeramic.Dimension.End.Row; stockC++)
                    {
                        var cellstockC = worksheetCeramic.Cells[stockC, 15];

                        if (cellstockC?.Value == null || !double.TryParse(cellstockC.Value.ToString(), out double stockvalueC) || stockvalueC == 0)  // all cases invalid
                        {
                            //Console.WriteLine($"cell null at ({stockC}, 15)  skip ceramic lot {worksheetCeramic.Cells[stockC, 3].Value}");

                            // delete that worksheet if the lot plug name skipping
                            string deletename = worksheetCeramic.Cells[stockC, 3].Value.ToString();

                            var deleteWS = packageMatrix.Workbook.Worksheets
                             .Where(ws => ws.Name.EndsWith("_" + deletename))
                             .ToList();
                            foreach (var deletesheet in deleteWS)
                            {
                                packageMatrix.Workbook.Worksheets.Delete(deletesheet);
                                //Console.WriteLine($"ceramic lot worksheet {worksheetCeramic.Cells[stockC, 3].Value} deleted");
                            }
                            copyceramic++;
                            continue;
                        }
                        //// only process if stock # 0, empty
                        // stock # 0
                        if (copyceramic <= worksheetCeramic.Dimension.End.Row)
                        {
                            var sourceceramic = worksheetCeramic.Cells[copyceramic, 3].Value;
                            if (sourceceramic == null)
                            {
                                //Console.WriteLine($"Source ceramic is null at copyceramic ={copyceramic}. Skipping.");
                                continue;
                            }
                            string prefixCERAMIC = sourceceramic.ToString();   // got prefix sourceceramic

                            // find target sheets to use that same prefix 
                            var ALLPLUG = packageMatrix.Workbook.Worksheets;
                            var targets = ALLPLUG
                                .Where(ws => ws.Name.EndsWith("_" + prefixCERAMIC))
                                .ToList();
                            //Console.WriteLine($"Prefix ceramic is : {prefixCERAMIC}");

                            foreach (var targetsheet in targets)    // in target sheet need to copy 
                            {
                                //Console.WriteLine($" Found ceramic sheet: {targetsheet.Name}");

                                // range data need to copy
                                for (int i = 0; i < 11; i++)
                                {
                                    var sourcevalue = worksheetCeramic.Cells[copyceramic, 4 + i].Value;     // from 1 source worksheetCeramic
                                    var targetCell = targetsheet.Cells[10, 4 + i];
                                    targetCell.Value = sourcevalue;
                                    targetCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    targetCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                                }
                            }

                        }
                        copyceramic++;
                    }





                    // -------------------------CALCULATE CELLS ----------------------------------

                    double ParseScientificNumber(string sciNumber)
                    {
                        if (string.IsNullOrEmpty(sciNumber))
                            return 0;

                        // Xử lý chuỗi dạng "1410e6"
                        if (sciNumber.Contains("e") || sciNumber.Contains("E"))
                        {
                            // Tách phần cơ số và số mũ
                            char[] separators = new char[] { 'e', 'E' };
                            string[] parts = sciNumber.Split(separators, 2);

                            if (parts.Length == 2)
                            {
                                if (double.TryParse(parts[0], out double baseValue) &&
                                    double.TryParse(parts[1], out double exponent))
                                {
                                    return baseValue * Math.Pow(10, exponent);
                                }
                            }
                        }

                        // Xử lý chuỗi số thông thường
                        if (double.TryParse(sciNumber, out double result))
                            return result;

                        // Trả về 0 nếu không thể chuyển đổi
                        return 0;
                    }
                    int startROW = 11;
                    int startCOL = 4;
                    int unit = 11;

                    var processsheet = packageMatrix.Workbook.Worksheets
                    .Where(ws => ws.Name.Contains("_"))
                    .ToList();
                    foreach (var process in processsheet)
                    {


                        //Ceramic - Plug
                        int countpass = 0;
                        int countfail = 0;
                        int total = 121;

                        for (int ROW = startROW; ROW < startROW + unit; ROW++)
                        {
                            for (int COL = startCOL; COL < startCOL + unit; COL++)
                            {
                                var cell = process.Cells[ROW, COL];
                                var big = process.Cells[10, COL].Value?.ToString();
                                var small = process.Cells[ROW, 3].Value?.ToString();
                                double bigvalue = ParseScientificNumber(big);
                                double smallvalue = ParseScientificNumber(small);
                                double cellvalue = bigvalue - smallvalue;
                                //return to string
                                string result = cellvalue.ToString("F5");
                                cell.Value = result;

                                // -----------------CHECK RANGE ---------------------
                                double epsilon = 1e-6;
                                if (cellvalue < 0.0015 - epsilon || cellvalue > 0.0035 + epsilon)
                                {
                                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                    countfail++;
                                }

                                else
                                {
                                    cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);
                                    countpass++;
                                }
                            }
                        }

                        // -------------------------------RESULT -------------------------
                        process.Cells["A1"].Value = "PASS";
                        process.Cells["A2"].Value = countpass;
                        var countpasscell = process.Cells["A2"];
                        countpasscell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        countpasscell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);

                        process.Cells["B1"].Value = "% PASS";
                        double percentpass = (double)countpass / total * 100;
                        process.Cells["B2"].Value = percentpass.ToString("F2") + "%";
                        var percentpasscell = process.Cells["B2"];
                        percentpasscell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        percentpasscell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);


                        process.Cells["C1"].Value = "FAIL";
                        process.Cells["C2"].Value = countfail;
                        var countfailcell = process.Cells["C2"];
                        countfailcell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        countfailcell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);

                        process.Cells["D1"].Value = "% FAIL";
                        double percentfail = (double)countfail / total * 100;
                        process.Cells["D2"].Value = percentfail.ToString("F2") + "%";
                        var percentfailcell = process.Cells["D2"];
                        percentfailcell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        percentfailcell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);

                    }

                    // ----------------------------MAIN TABLE PLUG ---------------------------------

                    int MPstartrow = 11;  // startrow at MAIN plug table

                    int plugMcopy = 11;   // lot plug scan in worksheetPlug

                    for (int stockP = 11; stockP <= worksheetPlug.Dimension.End.Row; stockP++)   // stockP scan top to bottom  
                    {
                        var cellstockP = worksheetPlug.Cells[stockP, 15];

                        if (cellstockP?.Value == null || !double.TryParse(cellstockP.Value.ToString(), out double stockvalue) || stockvalue == 0)  // all cases invalid
                        {
                            plugMcopy++;
                            continue;
                        }
                        //// only process if stock # 0, empty
                        // stock: 3000 # 0
                        if (plugMcopy <= worksheetPlug.Dimension.End.Row)
                        {
                            var sourceplug = worksheetPlug.Cells[plugMcopy, 3].Value;   // A,B,C,... (11,3)
                            if (sourceplug == null)
                            {
                                //Console.WriteLine($"Source plug is null at copyplug={copyplug}. Skipping.");
                                continue;
                            }
                            var MAINplug = worksheetMatrix.Cells[MPstartrow, 1];
                            MAINplug.Value = sourceplug;
                            MAINplug.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            MAINplug.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                            MPstartrow++;
                        }
                        plugMcopy++;
                    }



                    // ----------------------------MAIN TABLE CERAMIC

                    int MCstartcol = 2;        // startcol in MAIN ceramic table 

                    int ceramicMcopy = 11;     // lot ceramic scan in worksheetCeramic

                    for (int stockC = 11; stockC <= worksheetCeramic.Dimension.End.Row; stockC++)
                    {
                        var cellstockC = worksheetCeramic.Cells[stockC, 15];
                        if (cellstockC?.Value == null || !double.TryParse(cellstockC.Value.ToString(), out double stockvalue) || stockvalue == 0)  // all cases invalid
                        {
                            ceramicMcopy++;
                            continue;
                        }
                        //// only process if stock # 0, empty
                        // stock: 3000 # 0
                        if (ceramicMcopy <= worksheetCeramic.Dimension.End.Row)
                        {
                            var sourceceramic = worksheetCeramic.Cells[ceramicMcopy, 3].Value;   // M,N,P... (11,3)
                            if (sourceceramic == null)
                            {
                                //Console.WriteLine($"Source plug is null at copyplug={copyplug}. Skipping.");
                                continue;
                            }
                            var MAINceramic = worksheetMatrix.Cells[10, MCstartcol];
                            MAINceramic.Value = sourceceramic;
                            MAINceramic.Style.Fill.PatternType = ExcelFillStyle.Solid;
                            MAINceramic.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightBlue);
                            MCstartcol++;
                        }
                        ceramicMcopy++;
                    }



                    //-------------------------CELL PERCENT ----------------------------------------

                    var mainsheet = packageMatrix.Workbook.Worksheets
                    .Where(ws => ws.Name.Contains("_"))
                    .ToList();

                    foreach (var main in mainsheet)
                    {
                        var parts = main.Name.Split('_');
                        if (parts.Length != 2) continue; // Skip malformed names
                        string Plug = parts[0];
                        string Ceramic = parts[1];
                        var percent = main.Cells["B2"].Value;

                        // --------------FIND MATCHING PREFIX P-------------------
                        int targetRow = -1;
                        for (int row = 11; row <= worksheetMatrix.Dimension.End.Row; row++)
                        {
                            if (worksheetMatrix.Cells[row, 1].Text == Plug)
                            {
                                targetRow = row;               // detect row
                                break;
                            }
                        }
                        // -------------FIND MATCHING PREFIX C -------------------
                        int targetCol = -1;
                        for (int col = 2; col <= worksheetMatrix.Dimension.End.Column; col++)
                        {
                            if (worksheetMatrix.Cells[10, col].Text == Ceramic)
                            {
                                targetCol = col;               // detect column
                                break;
                            }
                        }
                        // --------------WRITE VALUE ----------------------
                        if (targetRow > 0 && targetCol > 0)                 // if both prefix detected
                        {
                            worksheetMatrix.Cells[targetRow, targetCol].Value = percent;
                        }
                    }


                    // -------------------------EVALUATE PERCENT --------------------------------

                    for (int row = 11; row <= worksheetMatrix.Dimension.End.Row; row++)
                    {
                        for (int col = 2; col <= worksheetMatrix.Dimension.End.Column; col++)
                        {
                            ExcelRange percentCell = null;
                            var cellValue = worksheetMatrix.Cells[row, col].Value?.ToString();

                            if (double.TryParse(cellValue?.Replace("%", ""), out double val))
                            {
                                percentCell = worksheetMatrix.Cells[row, col];
                                if (val >= 95)
                                {
                                    percentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    percentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);
                                }
                                if (val >= 85 && val < 95)
                                {
                                    percentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    percentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                                }
                                if (val < 85)
                                {
                                    percentCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    percentCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                }

                            }

                        }
                    }



                    //-------------------------------------------------------------
                    packageMatrix.SaveAs(new FileInfo(outputExcel)); // save file excel
                    UpdateStatus($"Processing output data completed, file saved to {outputExcel}");
                    MessageBox.Show($"Processing file Excel completed, file saved to {outputExcel} ", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    processedtextbox.Text = outputExcel;

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }



        }


        private void LoadFiles()
        {
            if (string.IsNullOrEmpty(outputExcel))
            {
                MessageBox.Show($"Please select the output first", "Reminder", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            UpdateStatus($"Loading Main Table...");
            System.Threading.Thread.Sleep(1500);

            try
            {
                 
                var fileInfo = new FileInfo(processedtextbox.Text);
                using (var package = new ExcelPackage(fileInfo))
                {
                    var worksheet = package.Workbook.Worksheets["MAIN"];
                    if (worksheet == null)
                    {

                        MessageBox.Show("(OUTPUT DATA) Can not find any worksheet named MAIN in file Excel.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    //// read data A2 (lot plug) and B2 (lot ceramic)
                    int numPlug = Convert.ToInt32(worksheet.Cells["A2"].Value);
                    int numCeramic = Convert.ToInt32(worksheet.Cells["B2"].Value);
                    
                    // clear the old data
                    datagridview.Rows.Clear();
                    datagridview.Columns.Clear();


                    //set up headers  ceramic N, P, O,...
                    int startcol = 2;
                    for (int col = 0; col < numCeramic; col++)
                    {
                        var Cname = worksheet.Cells[10, startcol].Value.ToString();
                        datagridview.Columns.Add($"{Cname}", $"{Cname}");
                        startcol++;
                    }
                    // Add rows and set row headers to A,B,C,D,...
                    int startrow = 11;
                    for (int row = 0; row < numPlug; row++)
                    {
                        var Pname = worksheet.Cells[startrow , 1].Value.ToString();
                        datagridview.Rows.Add(); // add empty row
                        datagridview.Rows[row].HeaderCell.Value = $"{Pname}";
                        startrow++;
                    }

                    ////Add data cells
                    int startRowInExcel = 11;
                    int startColInExcel = 2;
                    for (int Drow = startRowInExcel; Drow < startRowInExcel + numPlug; Drow++)
                    {
                        for (int Dcol = startColInExcel; Dcol < startColInExcel + numCeramic; Dcol++)
                        {
                            //object data = worksheet.Cells[Drow, Dcol].Value;
                            var value = worksheet.Cells[Drow, Dcol].Value?.ToString() ?? "";
                            int rowIndex = Drow - startRowInExcel;
                            int colIndex = Dcol - startColInExcel;
                            datagridview.Rows[rowIndex].Cells[colIndex].Value = value;

                        }
                    }
                }

                //------------------------------EVALUATE PERCENT CELL ---------------------------------

                int percentRow = -1;    
                int percentCol = -1;

                // Loop through each cell starting from [0,0]
                for (int row = 0; row < datagridview.RowCount; row++)
                {
                    for (int col = 0; col < datagridview.ColumnCount; col++)
                    {
                        var cellValue = datagridview.Rows[row].Cells[col].Value?.ToString();

                        if (double.TryParse(cellValue?.Replace("%", ""), out double val))
                        {
                            if (val >= 95)
                            {
                                percentRow = row;
                                percentCol = col;
                                datagridview.Rows[percentRow].Cells[percentCol].Style.BackColor = System.Drawing.Color.Green;
                            }
                            if (val >= 85 && val < 95)
                            {
                                percentRow = row;
                                percentCol = col;
                                datagridview.Rows[percentRow].Cells[percentCol].Style.BackColor = System.Drawing.Color.Yellow;
                            }
                            if (val < 85)
                            {
                                percentRow = row;
                                percentCol = col;
                                datagridview.Rows[percentRow].Cells[percentCol].Style.BackColor = System.Drawing.Color.Red;
                            }
                        }
                    }
                }





            }
            catch (Exception ex)
            {

                    MessageBox.Show($"Error loading: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
            }
        }



        private void CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                // Step 1: get 2 prefixs from cell position clicked
                // cell clicked, ignore rowheader and columnheader
                if (e.ColumnIndex >= 0 && e.RowIndex >= 0)
                {
                    // take the prefix of datagridview
                    string Pmain = datagridview.Rows[e.RowIndex].HeaderCell.Value?.ToString();
                    string Cmain = datagridview.Columns[e.ColumnIndex].HeaderText;

                    //take the prefix of every worksheets
                    var fileInfo = new FileInfo(processedtextbox.Text);
                    using (var packageMatrix = new ExcelPackage(fileInfo))
                    {
                        var mainsheet = packageMatrix.Workbook.Worksheets
                        .Where(ws => ws.Name.Contains("_"))
                        .ToList();
                        foreach (var main in mainsheet)
                        {
                            var parts = main.Name.Split('_');
                            if (parts.Length != 2) continue;
                            string Pdata = parts[0];
                            string Cdata = parts[1];

                            // compare prefix
                            if (Pdata == Pmain && Cdata == Cmain)         // detect worksheet matched
                            {
                                var matchedsheet = main;
                                //---------------------READ DATA FROM MATCHEDSHEET --------------------------

                                int popuppass = Convert.ToInt32(matchedsheet.Cells["A2"].Value);
                                int popupfail = Convert.ToInt32(matchedsheet.Cells["C2"].Value);
                                string popupPpass = Convert.ToString(matchedsheet.Cells["B2"].Value);
                                string popupPfail = Convert.ToString(matchedsheet.Cells["D2"].Value);

                                // original data from datagridview 
                                string columnName = datagridview.Columns[e.ColumnIndex].HeaderText;
                                string rowName = datagridview.Rows[e.RowIndex].HeaderCell.Value?.ToString() ?? $"Row {e.RowIndex + 1}";

                                // create form
                                Form popup = new Form();
                                popup.Text = $"Cell Detail [{columnName}, {rowName}] ";
                                popup.Size = new Size(900, 500);


                                FlowLayoutPanel panel = new FlowLayoutPanel();
                                panel.FlowDirection = FlowDirection.LeftToRight;
                                panel.Dock = DockStyle.Fill;


                                //--------------------------LABEL ------------------------------

                                // label 1: % PASS
                                Label label1 = new Label();
                                label1.Text = $"% PASS: {popupPpass} ";
                                label1.AutoSize = true;
                                label1.Font = new Font("Segoe UI", 14, FontStyle.Bold);

                                // label 2: PASS 
                                Label label2 = new Label();
                                label2.Text = $"PASS: {popuppass} ";
                                label2.AutoSize = true;
                                label2.Font = new Font("Segoe UI", 14, FontStyle.Bold);

                                // label 3: FAIL
                                Label label3 = new Label();
                                label3.Text = $"FAIL: {popupfail} ";
                                label3.AutoSize = true;
                                label3.Font = new Font("Segoe UI", 14, FontStyle.Bold);

                                //label 4: % FAIL
                                Label label4 = new Label();
                                label4.Text = $"% FAIL: {popupPfail} ";
                                label4.AutoSize = true;
                                label4.Font = new Font("Segoe UI", 14, FontStyle.Bold);


                                panel.Controls.Add(label1);
                                panel.Controls.Add(label2);
                                panel.Controls.Add(label3);
                                panel.Controls.Add(label4);

                               
                                DataGridView popupgrid = new DataGridView
                                {
                                    Dock = DockStyle.Fill,
                                    ReadOnly = true,
                                    ColumnCount = 2,
                                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
                                };

                                // Layout: TableLayoutPanel
                                TableLayoutPanel layout = new TableLayoutPanel();
                                layout.Dock = DockStyle.Fill;
                                layout.RowCount = 2;
                                layout.ColumnCount = 1;
                                layout.RowStyles.Add(new RowStyle(SizeType.AutoSize)); // For label
                                layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100)); // For grid
                                layout.Controls.Add(panel, 0, 0);              
                                layout.Controls.Add(popupgrid, 0, 1);

                                //---------------------------------------------
                                popupgrid.Rows.Clear();              
                                popupgrid.Columns.Clear();
                                


                                //--------------------ADD ROW AND COLUMN HEADER ----------------------
                                for (int col = 4; col <= matchedsheet.Dimension.End.Column; col++)
                                {
                                    var cellcolumn = matchedsheet.Cells[10, col];
                                    string columnpopup = cellcolumn.Text;
                                    string headerText = string.IsNullOrWhiteSpace(columnpopup) ? $"Col{col}" : columnpopup;

                                    // Only add column if the header cell has content
                                    if (!string.IsNullOrWhiteSpace(columnpopup))
                                    {
                                        popupgrid.Columns.Add($"col{col}", headerText);
                                    }
                                }



                                for (int row = 11; row <= matchedsheet.Dimension.End.Row; row++)
                                {
                                    var cellrow = matchedsheet.Cells[row, 3];  

                                    int newRowIndex = popupgrid.Rows.Add(); 

                                    popupgrid.Rows[newRowIndex].HeaderCell.Value = cellrow.Text; 
                                }

                                // ----------------------ADD CELLS --------------------------------


                                int startrow = 11;
                                int startcol = 4;
                               
                                for (int row = startrow; row < startrow + 11; row++)
                                {
                                    for (int col = startcol; col < startcol + 11; col++)
                                    {
                                        var cellmain = matchedsheet.Cells[row, col];   // position in excel
                                        int ROW = row - 11;
                                        int COL = col - 4;

                                        popupgrid.Rows[ROW].Cells[COL].Value = cellmain.Value;

                                        var cellcolor = cellmain.Style.Fill.BackgroundColor;
                                        if (cellcolor?.Rgb != null)  // Rgb is set only if the color is defined
                                        {
                                            var popupcolor = System.Drawing.ColorTranslator.FromHtml("#" + cellcolor.Rgb);
                                            popupgrid.Rows[ROW].Cells[COL].Style.BackColor = popupcolor;
                                        }
                                    }
                                }

                                //---------------------------------------------



                                // Add to form 
                                popup.StartPosition = FormStartPosition.Manual;
                                popup.Location = Cursor.Position;

                                //popup.Controls.Add(panel);
                                popup.Controls.Add(layout);

                                popup.ShowDialog();
                                break; // Exit loop once matched sheet is processed

                            }
                        }
                    }
                }
            }
            catch ( Exception ex )
            {
                UpdateStatus($"Error: {ex.Message}");
                MessageBox.Show($"Error Loading Table: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
        }


    }
}

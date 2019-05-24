using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MergeExcelTab
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnSelectSource_Click(object sender, EventArgs e)
        {
            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Excel|*.xlsx";
            DialogResult result = open.ShowDialog();
            if (result.Equals(DialogResult.OK))
            {
                txtSourcePath.Text = open.FileName;
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnSelectSavePath_Click(object sender, EventArgs e)
        {
            SaveFileDialog save = new SaveFileDialog();
            save.FileName = "merge";
            save.Filter = "Excel|*.xlsx";
            DialogResult result = save.ShowDialog();
            if (result.Equals(DialogResult.OK))
            {
                txtSavePath.Text = save.FileName;
            }
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            try
            {
                String source = txtSourcePath.Text.Trim();
                String result = txtSavePath.Text.Trim();
                if (source.Length == 0)
                {
                    MessageBox.Show("请选择合并的文件");
                    return;
                }
                if(result.Length == 0)
                {
                    MessageBox.Show("请选择保存的路径");
                    return;
                }
                List<List<CellEntity>> content = read();
                write(content);


            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }
        }

        private List<List<CellEntity>> read()
        {
            String source = txtSourcePath.Text.Trim();
            Aspose.Cells.Workbook excel = new Aspose.Cells.Workbook();
            excel.Open(source, Aspose.Cells.FileFormatType.Excel2007Xlsx);
            Worksheets sheets = excel.Worksheets;
            List<List<CellEntity>> contents = new List<List<CellEntity>>();
            for (int i = 0; i < sheets.Count; i++)
            {
                Worksheet sheet = sheets[i];
                int startRow = Int32.Parse(startRowSelect.Value.ToString());
                if (i == 0)
                {
                    startRow = 0;
                }
                int endRow = sheet.Cells.MaxDataRow; //结束行, 
                int colCount = sheet.Cells.Columns.Count;
                for (int rowIndex = startRow; rowIndex <= endRow; rowIndex++)
                {
                    List<CellEntity> row = new List<CellEntity>();
                    for (int colIndex = 0; colIndex < colCount; colIndex++)
                    {
                        Cell cell = sheet.Cells[rowIndex, colIndex];
                        Style style = cell.GetStyle();
                        String value = cell.HtmlString;
                        CellEntity cellEntity = new CellEntity();
                        cellEntity.Style = style;
                        cellEntity.Value = value;
                        cellEntity.Cell = cell;
                        //String content = ConvertObjectToString(sheet.Cells[rowIndex, colIndex].Value);
                        row.Add(cellEntity);
                    }
                    contents.Add(row);

                }
            }
            return contents;
        }

        private void write(List<List<CellEntity>> content)
        {
            try
            {
                String result = txtSavePath.Text.Trim();
                Aspose.Cells.Workbook excel = new Aspose.Cells.Workbook();
                Worksheet sheet = excel.Worksheets[0];

                for (int rowIndex = 0; rowIndex < content.Count; rowIndex++)
                {
                    List<CellEntity> row = content[rowIndex];
                    for (int colIndex = 0; colIndex < row.Count; colIndex++)
                    {
                        //Cell cell = sheet.Cells[currentRowIndex, colIndex];
                        // sheet.Cells.
                        CellEntity cellEntity = row[colIndex];
                        //sheet.Cells[currentRowIndex, colIndex].SetStyle(cellEntity.Style);

                        sheet.Cells[rowIndex, colIndex].Copy(cellEntity.Cell);
                        sheet.Cells[rowIndex, colIndex].PutValue(cellEntity.Cell.Value);
                        //sheet.Cells[currentRowIndex].Formula
                        //sheet.Cells[currentRowIndex, colIndex] = row[colIndex];
                        //sheet.Cells[currentRowIndex, colIndex].PutValue(row[colIndex]);
                    }


                }
                sheet.AutoFitColumns();
                excel.Save(result, FileFormatType.Excel2007Xlsx);
                MessageBox.Show("success");
            }
            catch (Exception)
            {

                throw;
            }
            
        }

        private string ConvertObjectToString(object value)
        {
            try
            {
                if (value != null)
                {
                    return value.ToString().Trim();
                }
            }
            catch { }
            return "";
        }
    }
}

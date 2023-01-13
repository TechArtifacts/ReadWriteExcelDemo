using OfficeOpenXml;

namespace ReadWriteExcelDemo
{
    public class ExcelHelper
    {
        public ExcelHelper()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        public IList<Employee> ReadEmployeeDataFromExcel(string filePath)
        {
            FileInfo fileInfo = new FileInfo(filePath);

            var employees = new List<Employee>();

            using (var package = new ExcelPackage(fileInfo))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count

                //We will start from row 2 since row 1 is a header row.
                for (int row = 2; row <= rowCount; row++)
                {
                    employees.Add(new Employee()
                    {
                        Id = Convert.ToInt32(worksheet.Cells[$"A{row}"].Value),
                        FirstName = worksheet.Cells[$"B{row}"].Value?.ToString(),
                        LastName = worksheet.Cells[$"C{row}"].Value?.ToString(),
                        MobileNumber = worksheet.Cells[$"D{row}"].Value?.ToString()
                    });
                }
            }
            return employees;
        }

        public void AddEmployeeData(string filePath, Employee employee)
        {
            FileInfo fileInfo = new FileInfo(filePath);

            using (var package = new ExcelPackage(fileInfo))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;     //get row count                

                //Add new data to the excel row
                worksheet.Cells[$"A{rowCount + 1}"].Value = employee.Id;
                worksheet.Cells[$"B{rowCount + 1}"].Value = employee.FirstName;
                worksheet.Cells[$"C{rowCount + 1}"].Value = employee.LastName;
                worksheet.Cells[$"D{rowCount + 1}"].Value = employee.MobileNumber;

                package.Save();
            }
        }
    }
}

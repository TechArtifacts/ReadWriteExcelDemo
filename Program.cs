
using ReadWriteExcelDemo;

var excelHelper = new ExcelHelper();

Console.WriteLine("Writing new employee");
//Add new employee
excelHelper.AddEmployeeData("C:\\TestData\\Employees.xlsx", new Employee()
{
    Id = 5,
    FirstName = "Suresh",
    LastName = "Raina",
    MobileNumber = "91-1111111115"
});

Console.WriteLine("Employee Data Saved Successfully!");

Console.WriteLine();

Console.WriteLine("Reading Employee Data");
var employees = excelHelper.ReadEmployeeDataFromExcel("C:\\TestData\\Employees.xlsx");

foreach (var employee in employees)
{
    Console.WriteLine($"Id: {employee.Id}, FirstName: {employee.FirstName}, LastName: {employee.LastName}, Mobile: {employee.MobileNumber}");
}

Console.Read();



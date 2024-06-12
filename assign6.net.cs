/******************************************************************************

                            Online C# Compiler.
                Code, Compile, Run and Debug C# program online.
Write your code in this editor and press "Run" button to execute it.

*******************************************************************************/

using System;
using System.Collections.Generic;
using System.IO;
using OfficeOpenXml; // You need to install EPPlus NuGet package for Excel manipulation

public abstract class BaseEntity
{
    public int Id { get; set; }
    // Other common properties if needed
}

public class EmployeeBasicDetails : BaseEntity
{
    public string Salutory { get; set; }
    public string FirstName { get; set; }
    public string MiddleName { get; set; }
    public string LastName { get; set; }
    public string NickName { get; set; }
    public string Email { get; set; }
    public string Mobile { get; set; }
    public string EmployeeID { get; set; }
    public string Role { get; set; }
    public string ReportingManagerUId { get; set; }
    public string ReportingManagerName { get; set; }
    public Address Address { get; set; }
}

public class EmployeeAdditionalDetails : BaseEntity
{
    public string EmployeeBasicDetailsUId { get; set; }
    public string AlternateEmail { get; set; }
    public string AlternateMobile { get; set; }
    public WorkInfo_ WorkInformation { get; set; }
    public PersonalDetails_ PersonalDetails { get; set; }
    public IdentityInfo_ IdentityInformation { get; set; }
}

public class WorkInfo_
{
    public string DesignationName { get; set; }
    public string DepartmentName { get; set; }
    public string LocationName { get; set; }
    public string EmployeeStatus { get; set; }
    public string SourceOfHire { get; set; }
    public DateTime DateOfJoining { get; set; }
}

public class PersonalDetails_
{
    public DateTime DateOfBirth { get; set; }
    public string Age { get; set; }
    public string Gender { get; set; }
    public string Religion { get; set; }
    public string Caste { get; set; }
    public string MaritalStatus { get; set; }
    public string BloodGroup { get; set; }
    public string Height { get; set; }
    public string Weight { get; set; }
}

public class IdentityInfo_
{
    public string PAN { get; set; }
    public string Aadhar { get; set; }
    public string Nationality { get; set; }
    public string PassportNumber { get; set; }
    public string PFNumber { get; set; }
}

public class Address
{
    // Define properties for Address entity if needed
}

public class EmployeeService
{
    // Implement CRUD operations for EmployeeBasicDetails and EmployeeAdditionalDetails

    public void ImportFromExcel(string filePath)
    {
        using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
            int rowCount = worksheet.Dimension.Rows;

            List<EmployeeBasicDetails> employees = new List<EmployeeBasicDetails>();

            for (int row = 2; row <= rowCount; row++) // Assuming first row is header
            {
                EmployeeBasicDetails employee = new EmployeeBasicDetails
                {
                    FirstName = worksheet.Cells[row, 2].Value.ToString(),
                    LastName = worksheet.Cells[row, 3].Value.ToString(),
                    Email = worksheet.Cells[row, 4].Value.ToString(),
                    Mobile = worksheet.Cells[row, 5].Value.ToString(),
                    ReportingManagerName = worksheet.Cells[row, 6].Value.ToString(),
                    // You need to parse dates appropriately
                };

                employees.Add(employee);
            }

            // Process imported data as needed
        }
    }

    public void ExportToExcel(List<EmployeeBasicDetails> employees, string filePath)
    {
        using (ExcelPackage package = new ExcelPackage())
        {
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Employees");

            // Add headers
            worksheet.Cells[1, 1].Value = "Sr.No";
            worksheet.Cells[1, 2].Value = "First Name";
            worksheet.Cells[1, 3].Value = "Last Name";
            worksheet.Cells[1, 4].Value = "Email";
            worksheet.Cells[1, 5].Value = "Phone No";
            worksheet.Cells[1, 6].Value = "Reporting Manager Name";
            worksheet.Cells[1, 7].Value = "Date Of Birth";
            worksheet.Cells[1, 8].Value = "Date of Joining";

            // Fill data
            for (int i = 0; i < employees.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = i + 1;
                worksheet.Cells[i + 2, 2].Value = employees[i].FirstName;
                worksheet.Cells[i + 2, 3].Value = employees[i].LastName;
                worksheet.Cells[i + 2, 4].Value = employees[i].Email;
                worksheet.Cells[i + 2, 5].Value = employees[i].Mobile;
                worksheet.Cells[i + 2, 6].Value = employees[i].ReportingManagerName;
                // Fill other columns accordingly
            }

            package.SaveAs(new FileInfo(filePath));
        }
    }
}

class Program
{
    static void Main(string[] args)
    {
        // Usage example
        EmployeeService employeeService = new EmployeeService();
        employeeService.ImportFromExcel("input.xlsx"); // Provide the path of input Excel file
        // Perform CRUD operations as needed
        // Export data to Excel
        // employeeService.ExportToExcel(employees, "output.xlsx"); // Provide the path for output Excel file
    }
}

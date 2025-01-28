using OfficeOpenXml;
using System.Diagnostics;
using System.Reflection;

namespace Reflection_Task
{
    public class Employee
    {
        int Id { get; set; }
        string Name { get; set; }
        decimal Salary { get; set; }
        string dept {  get; set; }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;


            Console.WriteLine(CheckExcelFile("test.xlsx", typeof(Employee)) ? "True" : "False");
            
        }

        /// <summary>
        /// Checks whether the column headers in the first row of an Excel file match the property names of a given type.
        /// </summary>
        /// <param name="filePath">The path to the Excel file to be checked.</param>
        /// <param name="t">The type whose property names will be compared against the Excel column headers.</param>
        /// <returns>
        /// <c>true</c> if all column headers in the Excel file match the property names of the given type (case-insensitive and trimmed);
        /// otherwise, <c>false</c>.
        /// </returns>
        /// <exception cref="ArgumentNullException">Thrown if <paramref name="filePath"/> or <paramref name="t"/> is null.</exception>
        /// <exception cref="FileNotFoundException">Thrown if the Excel file at <paramref name="filePath"/> does not exist.</exception>
        /// <remarks>
        /// This method reads the first row of the first worksheet in the Excel file and compares each cell value (column header)
        /// to the property names of the specified type. The comparison is case-insensitive and ignores leading/trailing whitespace.
        /// If any column header does not match a property name, the method returns <c>false</c>.
        /// </remarks>
        static bool CheckExcelFile(string filePath, Type t)
        {
            
            //// Get all class properties
            List<string> propertiesNames = t.GetProperties(BindingFlags.Static | BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic).Select(x => x.Name.ToLower().Trim()).ToList();

            //// Check the name of the heading row
            using (var package = new ExcelPackage(filePath))
            {
                // Select the first worksheet
                using (var sheet = package.Workbook.Worksheets[0])
                {
                    // Iterate through the cells of the first column
                    for (int i = 1; i <= sheet.Dimension.Columns; i++)
                    {
                        string cellValue = sheet.Cells[1, i].Value.ToString().ToLower().Trim();

                        if (string.IsNullOrEmpty(cellValue))
                            continue;
                        else if (!propertiesNames.Contains(cellValue))
                            return false;
                    }
                }
            }

            return true;
        }
    }
}

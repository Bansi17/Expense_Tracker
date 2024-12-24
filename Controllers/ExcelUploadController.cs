using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.IO;
using System.Threading.Tasks;

namespace Expense_Tracker.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelUploadController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public ExcelUploadController(IConfiguration configuration)
        {
            _configuration = configuration;
        }



        [HttpPost("upload")]
        public async Task<IActionResult> UploadFile(IFormFile file)
        {
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                stream.Position = 0;

                using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.First();
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                    var rowCount = sheetData.Elements<Row>().Count();

                    var connectionString = _configuration.GetConnectionString("DefaultConnection");

                    using (var connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        int rowIndex = 1; // Skipping header row
                        foreach (var row in sheetData.Elements<Row>().Skip(1))
                        {
                            

                            var columns = new string[24];
                            int colIndex = 0;

                            foreach (var cell in row.Elements<Cell>())
                            {
                                var cellValue = GetCellValue(cell, workbookPart);
                                if (colIndex < 24)
                                {
                                    columns[colIndex++] = cellValue;
                                }
                            }

                            if (columns.Length == 24)
                            {
                                var commandText = @"
                                    INSERT INTO imageDetails (
                                        Planned, [GIS Code], [Advertisement Type], 
                                        [Planned Area], [Brand Name], [SubBrand Name],
                                        Name, [No.Of WP], [Plan Date], IMEI, 
                                        [DO Number], [Plan Type], [End Date],
                                        [HUL Person], Status, [State Name], 
                                        [State Code], [District Name], [District Code], 
                                        [Tehsil Name], [Tehsil Code], [Village Name], 
                                        [Village Code], [Village Population]
                                    ) VALUES (
                                        @Planned, @GISCode, @AdvertisementType, 
                                        @PlannedArea, @BrandName, @SubBrandName, 
                                        @Name, @NoOfWP, @PlanDate, @IMEI, 
                                        @DONumber, @PlanType, @EndDate, 
                                        @HULPerson, @Status, @StateName, 
                                        @StateCode, @DistrictName, @DistrictCode, 
                                        @TehsilName, @TehsilCode, @VillageName, 
                                        @VillageCode, @VillagePopulation
                                    )";

                                using (var command = new SqlCommand(commandText, connection))
                                {
                                    command.Parameters.AddWithValue("@Planned", columns[0]);
                                    command.Parameters.AddWithValue("@GISCode", columns[1]);
                                    command.Parameters.AddWithValue("@AdvertisementType", columns[2]);
                                    command.Parameters.AddWithValue("@PlannedArea", columns[3]);
                                    command.Parameters.AddWithValue("@BrandName", columns[4]);
                                    command.Parameters.AddWithValue("@SubBrandName", columns[5]);
                                    command.Parameters.AddWithValue("@Name", columns[6]);
                                    command.Parameters.AddWithValue("@NoOfWP", columns[7]);
                                    command.Parameters.AddWithValue("@PlanDate", columns[8]);
                                    command.Parameters.AddWithValue("@IMEI", columns[9]);
                                    command.Parameters.AddWithValue("@DONumber", columns[10]);
                                    command.Parameters.AddWithValue("@PlanType", columns[11]);
                                    command.Parameters.AddWithValue("@EndDate", columns[12]);
                                    command.Parameters.AddWithValue("@HULPerson", columns[13]);
                                    command.Parameters.AddWithValue("@Status", columns[14]);
                                    command.Parameters.AddWithValue("@StateName", columns[15]);
                                    command.Parameters.AddWithValue("@StateCode", columns[16]);
                                    command.Parameters.AddWithValue("@DistrictName", columns[17]);
                                    command.Parameters.AddWithValue("@DistrictCode", columns[18]);
                                    command.Parameters.AddWithValue("@TehsilName", columns[19]);
                                    command.Parameters.AddWithValue("@TehsilCode", columns[20]);
                                    command.Parameters.AddWithValue("@VillageName", columns[21]);
                                    command.Parameters.AddWithValue("@VillageCode", columns[22]);

                                    //command.Parameters.AddWithValue("@VillagePopulation", string.IsNullOrEmpty(columns[23]) ? DBNull.Value : columns[23]);
                                    command.Parameters.AddWithValue("@VillagePopulation", string.IsNullOrEmpty(columns[23]) ? DBNull.Value : columns[23]);


                                    try
                                    {
                                        await command.ExecuteNonQueryAsync();
                                        Console.WriteLine("Row inserted successfully.");
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error inserting row: {ex.Message}");
                                    }
                                }
                            }
                        }
                        connection.Close();
                    }
                }

                return Ok("File uploaded and data inserted successfully.");
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }



        // New GET API to retrieve all data from the imageDetails table
        [HttpGet("get-all-details")]
        public IActionResult GetAllData()
        {
            var connectionString = _configuration.GetConnectionString("DefaultConnection");

            try
            {
                var dataList = new List<object>();

                using (var connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    var query = "SELECT * FROM imageDetails";
                    using (var command = new SqlCommand(query, connection))
                    {
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var record = new
                                {
                                    ID = reader["ID"],
                                    Planned = reader["Planned"],
                                    GISCode = reader["GIS Code"],
                                    AdvertisementType = reader["Advertisement Type"],
                                    PlannedArea = reader["Planned Area"],
                                    BrandName = reader["Brand Name"],
                                    SubBrandName = reader["SubBrand Name"],
                                    Name = reader["Name"],
                                    NoOfWP = reader["No.Of WP"],
                                    PlanDate = reader["Plan Date"],
                                    IMEI = reader["IMEI"],
                                    DONumber = reader["DO Number"],
                                    PlanType = reader["Plan Type"],
                                    EndDate = reader["End Date"],
                                    HULPerson = reader["HUL Person"],
                                    Status = reader["Status"],
                                    StateName = reader["State Name"],
                                    StateCode = reader["State Code"],
                                    DistrictName = reader["District Name"],
                                    DistrictCode = reader["District Code"],
                                    TehsilName = reader["Tehsil Name"],
                                    TehsilCode = reader["Tehsil Code"],
                                    VillageName = reader["Village Name"],
                                    VillageCode = reader["Village Code"],
                                    VillagePopulation = reader["Village Population"]
                                };

                                dataList.Add(record);
                            }
                        }
                    }
                }

                return Ok(dataList);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            var value = cell.CellValue?.Text;

            if (cell.DataType == null || cell.DataType.Value != CellValues.SharedString)
            {
                return value;
            }

            var sharedStringTablePart = workbookPart.SharedStringTablePart;
            if (sharedStringTablePart != null)
            {
                var sharedStringItem = sharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(int.Parse(value));
                return sharedStringItem.Text?.Text ?? sharedStringItem.InnerText;
            }

            return value;
        }
    }
}

using ExcelUploadgridApp.Data;
using ExcelUploadgridApp.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using System.Linq;
using EFCore.BulkExtensions;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Data;
using Microsoft.Data.SqlClient;
using CsvHelper.Configuration;
using CsvHelper;
using System.Globalization;
using OfficeOpenXml;
using DocumentFormat.OpenXml.InkML;
using OfficeOpenXml.Export.HtmlExport.StyleCollectors.StyleContracts;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using CellType = NPOI.SS.UserModel.CellType;

namespace ExcelUploadgridApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelUploadController : ControllerBase
    {
        private readonly AppDbContext _context;
        private readonly DbContextOptions<AppDbContext> _dbContextOptions;
        private readonly ILogger<ExcelUploadController1> _logger;
        public ExcelUploadController(AppDbContext context, DbContextOptions<AppDbContext> dbContextOptions, ILogger<ExcelUploadController1> logger)
        {
            _context = context;
            _dbContextOptions = dbContextOptions;
            _logger = logger;
        }

        [HttpGet("datasources")]
        public async Task<IActionResult> GetDatasources()
        {
            var datasources = await _context.ParentDetails
                                            .Where(d => d.IsActive)
                                            .Select(d => new { d.Id, d.DatasourceName })
                                            .ToListAsync();
            return Ok(datasources);
        }

        [HttpGet("columns/{datasourceId}")]
        public async Task<IActionResult> GetColumns(int datasourceId)
        {
            var columns = await _context.ColumnDetails
                                        .Where(c => c.ParentId == datasourceId)
                                        .Select(c => new { c.ColumnName, c.DataType })
                                        .ToListAsync();
            return Ok(columns);
        }

        [HttpPost("upload")]
        public async Task<IActionResult> Upload(IFormFile file, [FromQuery] int datasourceId)
        {
            var stopwatch = new Stopwatch();

            _logger.LogInformation("Starting to process the uploaded file.");

            stopwatch.Start();
            var columns = await _context.ColumnDetails
                                        .Where(c => c.ParentId == datasourceId)
                                        .Select(c => new
                                        {
                                            c.UserFriendlyName,
                                            c.ConstraintExpression,
                                            c.ColumnName
                                        })
                                        .ToListAsync();
            stopwatch.Stop();
            _logger.LogInformation($"Fetched columns from the database in {stopwatch.ElapsedMilliseconds} ms.");

            var excelColumns = new List<string>();
            var data = new ConcurrentBag<Dictionary<string, object>>();
            var validationResults = new ConcurrentBag<Dictionary<string, string>>();

            using (var stream = file.OpenReadStream())
            {
                var workbook = new XSSFWorkbook(stream);
                var sheet = workbook.GetSheetAt(0);

                if (sheet.PhysicalNumberOfRows == 0)
                {
                    return BadRequest(new { message = "The Excel file is empty" });
                }

                var headerRow = sheet.GetRow(0);
                for (int i = 0; i < headerRow.LastCellNum; i++)
                {
                    excelColumns.Add(headerRow.GetCell(i).ToString());
                }

                var missingColumns = columns.Select(c => c.UserFriendlyName).Except(excelColumns).ToList();
                if (missingColumns.Any())
                {
                    return BadRequest(new { message = "Excel file is missing required columns", missingColumns });
                }

                stopwatch.Restart();

                int batchSize = 1000;  // Define your batch size
                int totalRows = sheet.PhysicalNumberOfRows;

                for (int i = 1; i < totalRows; i += batchSize)
                {
                    var batchRows = new List<IRow>();
                    for (int j = i; j < i + batchSize && j < totalRows; j++)
                    {
                        batchRows.Add(sheet.GetRow(j));
                    }

                    var tasks = batchRows.Select(row => Task.Run(() =>
                    {
                        var rowData = new Dictionary<string, object>();
                        var rowValidation = new Dictionary<string, string>();

                        for (int colIndex = 0; colIndex < row.LastCellNum; colIndex++)
                        {
                            var cell = row.GetCell(colIndex);
                            var columnName = excelColumns[colIndex];
                            var cellValue = GetCellValue(cell);

                            rowData[columnName] = cellValue;

                            var column = columns.FirstOrDefault(c => c.UserFriendlyName == columnName);
                            if (column != null && !string.IsNullOrEmpty(column.ConstraintExpression))
                            {
                                // Example validation checks based on column.ConstraintExpression
                                ValidateCell(column.ConstraintExpression, cellValue, rowValidation, columnName);
                            }
                        }
                        data.Add(rowData);
                        validationResults.Add(rowValidation);
                    }));

                    await Task.WhenAll(tasks);
                }

                stopwatch.Stop();
                _logger.LogInformation($"Processed rows in {stopwatch.ElapsedMilliseconds} ms.");
            }

            return Ok(new { columns = excelColumns.Select(col => new { Header = col }), data, validationResults });
        }

        private void ValidateCell(string constraintExpression, string cellValue, Dictionary<string, string> rowValidation, string columnName)
        {
            // Simplified validation logic
            switch (constraintExpression)
            {
                case "value > 100":
                    if (int.TryParse(cellValue, out var intValue) && intValue <= 100)
                    {
                        rowValidation[columnName] = "Value must be greater than 100";
                    }
                    break;
                case "value <= DateTime.Now":
                    if (DateTime.TryParse(cellValue, out var dateValue) && dateValue > DateTime.Now)
                    {
                        rowValidation[columnName] = "Date must be in the past";
                    }
                    break;
                case "value == 0 || value == 1":
                    if (int.TryParse(cellValue, out var bitValue) && (bitValue != 0 && bitValue != 1))
                    {
                        rowValidation[columnName] = "Value must be 0 or 1";
                    }
                    break;
                case "value > 30":
                    if (int.TryParse(cellValue, out var ageValue) && ageValue <= 30)
                    {
                        rowValidation[columnName] = "Age must be greater than 30";
                    }
                    break;
               /* case "value != null && value.Trim() != ''":
                    if (string.IsNullOrEmpty(cellValue?.Trim()))
                    {
                        rowValidation[columnName] = "Value must not be empty";
                    }*/
                    break;
                // Add more validation checks as needed
                default:
                    break;
            }
        }

        private string GetCellValue(ICell cell)
        {
            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Numeric:
                    return cell.NumericCellValue.ToString();
                case CellType.Boolean:
                    return cell.BooleanCellValue.ToString();
                default:
                    return cell.ToString();
            }
        }

        [HttpPost("upload-csv")]
        public async Task<IActionResult> UploadCsv(IFormFile file, [FromQuery] int datasourceId)
        {
            var stopwatch = new Stopwatch();

            _logger.LogInformation("Starting to process the uploaded CSV file.");

            stopwatch.Start();
            var columns = await _context.ColumnDetails
                                        .Where(c => c.ParentId == datasourceId)
                                        .Select(c => new {
                                            c.UserFriendlyName,
                                            c.ConstraintExpression,
                                            c.ColumnName
                                        })
                                        .ToListAsync();
            stopwatch.Stop();
            _logger.LogInformation($"Fetched columns from the database in {stopwatch.ElapsedMilliseconds} ms.");

            var csvColumns = new List<string>();
            var data = new ConcurrentBag<Dictionary<string, object>>();
            var validationResults = new ConcurrentBag<Dictionary<string, string>>();

            stopwatch.Restart();
            using (var stream = file.OpenReadStream())
            {
                using (var reader = new StreamReader(stream))
                using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)))
                {
                    stopwatch.Stop();
                    _logger.LogInformation($"Opened CSV file in {stopwatch.ElapsedMilliseconds} ms.");

                    csv.Read();
                    csv.ReadHeader();
                    var headerRecord = csv.HeaderRecord;

                    if (headerRecord == null || !headerRecord.Any())
                    {
                        return BadRequest(new { message = "The CSV file is empty or does not have a header row" });
                    }

                    csvColumns.AddRange(headerRecord);

                    var missingColumns = columns.Select(c => c.UserFriendlyName).Except(csvColumns).ToList();
                    if (missingColumns.Any())
                    {
                        return BadRequest(new { message = "CSV file is missing required columns", missingColumns });
                    }

                    stopwatch.Restart();
                    var records = csv.GetRecords<dynamic>().ToList();
                    Parallel.ForEach(records, record =>
                    {
                        var rowData = new Dictionary<string, object>();
                        var rowValidation = new Dictionary<string, string>();

                        foreach (var column in columns)
                        {
                            var columnName = column.UserFriendlyName;
                            if (((IDictionary<string, object>)record).TryGetValue(columnName, out var cellValue))
                            {
                                rowData[columnName] = cellValue;

                                if (!string.IsNullOrEmpty(column.ConstraintExpression))
                                {
                                    // Example validation checks based on column.ConstraintExpression
                                    switch (column.ConstraintExpression)
                                    {
                                        case "value > 100":
                                            if (int.TryParse(cellValue?.ToString(), out var intValue) && intValue <= 100)
                                            {
                                                rowValidation[columnName] = "Value must be greater than 100";
                                            }
                                            break;
                                        case "value <= DateTime.Now":
                                            if (DateTime.TryParse(cellValue?.ToString(), out var dateValue) && dateValue > DateTime.Now)
                                            {
                                                rowValidation[columnName] = "Date must be in the past";
                                            }
                                            break;
                                        case "value == 0 || value == 1":
                                            if (int.TryParse(cellValue?.ToString(), out var bitValue) && (bitValue != 0 && bitValue != 1))
                                            {
                                                rowValidation[columnName] = "Value must be 0 or 1";
                                            }
                                            break;
                                        case "value > 30":
                                            if (int.TryParse(cellValue?.ToString(), out var ageValue) && ageValue <= 30)
                                            {
                                                rowValidation[columnName] = "Age must be greater than 30";
                                            }
                                            break;
                                        case "value != null && value.Trim() != ''":
                                            if (string.IsNullOrEmpty(cellValue?.ToString()?.Trim()))
                                            {
                                                rowValidation[columnName] = "Value must not be empty";
                                            }
                                            break;
                                        // Add more validation checks as needed
                                        default:
                                            break;
                                    }
                                }
                            }
                        }
                        data.Add(rowData);
                        validationResults.Add(rowValidation);
                    });
                    stopwatch.Stop();
                    _logger.LogInformation($"Processed rows in {stopwatch.ElapsedMilliseconds} ms.");
                }
            }

            return Ok(new { columns = csvColumns.Select(col => new { Header = col }), data, validationResults });
        }


        [HttpPost("saveX")]
        public async Task<IActionResult> SaveValuesX(int dataSourceId, [FromBody] List<Dictionary<string, object>> editedValues)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            _context.Database.SetCommandTimeout(300); // Timeout in seconds

            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                try
                {
                    var columnDetails = await _context.ColumnDetails
                                                       .Where(c => c.ParentId == dataSourceId)
                                                       .Select(c => new
                                                       {
                                                           c.UserFriendlyName,
                                                           c.ConstraintExpression,
                                                           c.ColumnName,
                                                           c.Id
                                                       })
                                                       .ToListAsync();

                    var uniqueColumnNames = editedValues.SelectMany(row => row.Keys).Distinct().ToList();

                    // Group edited values by column
                    var groupedValuesByColumn = uniqueColumnNames.ToDictionary(
                        columnName => columnName,
                        columnName => new List<(int rowId, object value)>()
                    );

                    foreach (var editedRow in editedValues)
                    {
                        int rowId = editedRow.ContainsKey("rowId") ? ((JsonElement)editedRow["rowId"]).GetInt32() : 0;

                        foreach (var columnName in uniqueColumnNames)
                        {
                            if (editedRow.TryGetValue(columnName, out object cellValue))
                            {
                                groupedValuesByColumn[columnName].Add((rowId, cellValue));
                            }
                        }
                    }

                    var rowIds = editedValues.Select(ev => ((JsonElement)ev["rowId"]).GetInt32()).Distinct().ToList();
                    var columnIds = columnDetails.Select(cd => cd.Id).ToList();
                    var existingValueDetails = await _context.ValueDetails
                                                              .Where(vd => rowIds.Contains(vd.RowId) && columnIds.Contains(vd.ColumnId))
                                                              .ToListAsync();

                    // Create DataTable for bulk operations
                    var dataTable = new DataTable();
                    dataTable.Columns.Add("ColumnId", typeof(int));
                    dataTable.Columns.Add("RowId", typeof(int));
                    dataTable.Columns.Add("Value", typeof(string));

                    // Populate DataTable
                    foreach (var column in columnDetails)
                    {
                        if (groupedValuesByColumn.TryGetValue(column.ColumnName.ToLower(), out var valuesForColumn))
                        {
                            foreach (var (rowId, cellValue) in valuesForColumn)
                            {
                                var valueDetail = existingValueDetails
                                    .FirstOrDefault(vd => vd.RowId == rowId && vd.ColumnId == column.Id);

                                if (valueDetail != null)
                                {
                                    // Update existing row
                                    valueDetail.Value = cellValue.ToString();
                                }
                                else
                                {
                                    // Add new row to DataTable
                                    var newRow = dataTable.NewRow();
                                    newRow["ColumnId"] = column.Id;
                                    newRow["RowId"] = rowId;
                                    newRow["Value"] = cellValue.ToString();
                                    dataTable.Rows.Add(newRow);
                                }
                            }
                        }
                    }

                    // Perform bulk update
                    if (existingValueDetails.Any())
                    {
                        await _context.BulkUpdateAsync(existingValueDetails);
                    }

                    // Perform bulk insert using SqlBulkCopy
                    if (dataTable.Rows.Count > 0)
                    {
                        using (var sqlBulkCopy = new SqlBulkCopy(_context.Database.GetConnectionString(), SqlBulkCopyOptions.KeepIdentity | SqlBulkCopyOptions.KeepNulls))
                        {
                            sqlBulkCopy.DestinationTableName = "ValueDetails";
                            sqlBulkCopy.ColumnMappings.Add("ColumnId", "ColumnId");
                            sqlBulkCopy.ColumnMappings.Add("RowId", "RowId");
                            sqlBulkCopy.ColumnMappings.Add("Value", "Value");

                            await sqlBulkCopy.WriteToServerAsync(dataTable);
                        }
                    }

                    await transaction.CommitAsync();

                    stopwatch.Stop();
                    _logger.LogInformation($"Save operation completed in {stopwatch.ElapsedMilliseconds} ms.");

                    return Ok(new { message = "Values saved successfully" });
                }
                catch (Exception ex)
                {
                    await transaction.RollbackAsync();
                    stopwatch.Stop();
                    _logger.LogError($"Save operation failed after {stopwatch.ElapsedMilliseconds} ms. Error: {ex.Message}");

                    return StatusCode(500, new { message = "An error occurred while saving values", error = ex.Message });
                }
            }
        }




        [HttpPost("save")]
        public async Task<IActionResult> SaveValues(int dataSourceId, [FromBody] List<Dictionary<string, object>> editedValues)
        {
            var stopwatch = new Stopwatch();
            stopwatch.Start();

            // Increase the command timeout
            _context.Database.SetCommandTimeout(300); // Timeout in seconds

            using (var transaction = await _context.Database.BeginTransactionAsync())
            {
                try
                {
                    // Fetch column details related to the provided dataSourceId
                    var columnDetails = await _context.ColumnDetails
                                                       .Where(c => c.ParentId == dataSourceId)
                                                       .Select(c => new
                                                       {
                                                           c.UserFriendlyName,
                                                           c.ColumnName,
                                                           c.Id
                                                       })
                                                       .ToListAsync();

                    var uniqueColumnNames = editedValues.SelectMany(row => row.Keys).Distinct().ToList();

                    // Group edited values by column
                    var groupedValuesByColumn = uniqueColumnNames.ToDictionary(
                        columnName => columnName,
                        columnName => new List<(int rowId, object value)>()
                    );

                    foreach (var editedRow in editedValues)
                    {
                        int rowId = editedRow.ContainsKey("rowId") ? ((JsonElement)editedRow["rowId"]).GetInt32() : 0;

                        foreach (var columnName in uniqueColumnNames)
                        {
                            if (editedRow.TryGetValue(columnName, out object cellValue))
                            {
                                groupedValuesByColumn[columnName].Add((rowId, cellValue));
                            }
                        }
                    }

                    var rowIds = editedValues.Select(ev => ((JsonElement)ev["rowId"]).GetInt32()).Distinct().ToList();
                    var columnIds = columnDetails.Select(cd => cd.Id).ToList();
                    var existingValueDetails = await _context.ValueDetails
                                                              .Where(vd => rowIds.Contains(vd.RowId) && columnIds.Contains(vd.ColumnId))
                                                              .ToListAsync();

                    var valueDetailsToUpdate = new List<ValueDetail>();
                    var valueDetailsToInsert = new List<ValueDetail>();

                    var tasks = columnDetails
                        .Where(column => groupedValuesByColumn.ContainsKey(column.ColumnName.ToLower()))
                        .Select(async column =>
                        {
                            var valuesForColumn = groupedValuesByColumn[column.ColumnName.ToLower()];

                            foreach (var (rowId, cellValue) in valuesForColumn)
                            {
                                var valueDetail = existingValueDetails
                                    .FirstOrDefault(vd => vd.RowId == rowId && vd.ColumnId == column.Id);

                                if (valueDetail != null)
                                {
                                    valueDetail.Value = cellValue.ToString();
                                    valueDetailsToUpdate.Add(valueDetail);
                                }
                                else
                                {
                                    valueDetailsToInsert.Add(new ValueDetail
                                    {
                                        ColumnId = column.Id,
                                        RowId = rowId,
                                        Value = cellValue.ToString()
                                    });
                                }
                            }
                        });

                    await Task.WhenAll(tasks);

                    // Process updates and inserts in batches
                    const int batchSize = 1000;

                    for (int i = 0; i < valueDetailsToUpdate.Count; i += batchSize)
                    {
                        var batch = valueDetailsToUpdate.Skip(i).Take(batchSize).ToList();
                        await _context.BulkUpdateAsync(batch);
                    }

                    for (int i = 0; i < valueDetailsToInsert.Count; i += batchSize)
                    {
                        var batch = valueDetailsToInsert.Skip(i).Take(batchSize).ToList();
                        await _context.BulkInsertAsync(batch);
                    }

                    await transaction.CommitAsync();

                    stopwatch.Stop();
                    _logger.LogInformation($"Save operation completed in {stopwatch.ElapsedMilliseconds} ms.");

                    return Ok(new { message = "Values saved successfully" });
                }
                catch (Exception ex)
                {
                    await transaction.RollbackAsync();
                    stopwatch.Stop();
                    _logger.LogError($"Save operation failed after {stopwatch.ElapsedMilliseconds} ms. Error: {ex.Message}");

                    return StatusCode(500, new { message = "An error occurred while saving values", error = ex.Message });
                }
            }
        }

        [HttpPost("upload-epplus")]
        public async Task<IActionResult> UploadEPPlus(IFormFile file, [FromQuery] int datasourceId)
        {
            var stopwatch = new Stopwatch();
            _logger.LogInformation("Starting to process the uploaded file.");

            stopwatch.Start();
            var columns = await _context.ColumnDetails
                                        .Where(c => c.ParentId == datasourceId)
                                        .Select(c => new
                                        {
                                            c.UserFriendlyName,
                                            c.ConstraintExpression,
                                            c.ColumnName
                                        })
                                        .ToListAsync();
            stopwatch.Stop();
            _logger.LogInformation($"Fetched columns from the database in {stopwatch.ElapsedMilliseconds} ms.");

            var excelColumns = new List<string>();
            var data = new ConcurrentBag<Dictionary<string, object>>();
            var validationResults = new ConcurrentBag<Dictionary<string, string>>();

            stopwatch.Restart();
            using (var stream = file.OpenReadStream())
            {
                using (var package = new ExcelPackage(stream))
                {
                    var worksheet = package.Workbook.Worksheets.FirstOrDefault();
                    if (worksheet == null)
                    {
                        return BadRequest(new { message = "The Excel file is empty" });
                    }

                    var rowCount = worksheet.Dimension.Rows;
                    var colCount = worksheet.Dimension.Columns;

                    // Reading header row
                    for (var col = 1; col <= colCount; col++)
                    {
                        var headerValue = worksheet.Cells[1, col].Text;
                        excelColumns.Add(headerValue);
                    }

                    var missingColumns = columns.Select(c => c.UserFriendlyName).Except(excelColumns).ToList();
                    if (missingColumns.Any())
                    {
                        return BadRequest(new { message = "Excel file is missing required columns", missingColumns });
                    }

                    stopwatch.Restart();
                    Parallel.For(2, rowCount + 1, row =>
                    {
                        var rowData = new Dictionary<string, object>();
                        var rowValidation = new Dictionary<string, string>();

                        for (var col = 1; col <= colCount; col++)
                        {
                            var columnName = excelColumns[col - 1];
                            var cellValue = worksheet.Cells[row, col].Text;

                            rowData[columnName] = cellValue;

                            var column = columns.FirstOrDefault(c => c.UserFriendlyName == columnName);
                            if (column != null && !string.IsNullOrEmpty(column.ConstraintExpression))
                            {
                                // Example validation checks based on column.ConstraintExpression
                                switch (column.ConstraintExpression)
                                {
                                    case "value > 100":
                                        if (int.TryParse(cellValue, out var intValue) && intValue <= 100)
                                        {
                                            rowValidation[columnName] = "Value must be greater than 100";
                                        }
                                        break;
                                    case "value <= DateTime.Now":
                                        if (DateTime.TryParse(cellValue, out var dateValue) && dateValue > DateTime.Now)
                                        {
                                            rowValidation[columnName] = "Date must be in the past";
                                        }
                                        break;
                                    case "value == 0 || value == 1":
                                        if (int.TryParse(cellValue, out var bitValue) && (bitValue != 0 && bitValue != 1))
                                        {
                                            rowValidation[columnName] = "Value must be 0 or 1";
                                        }
                                        break;
                                    case "value > 30":
                                        if (int.TryParse(cellValue, out var ageValue) && ageValue <= 30)
                                        {
                                            rowValidation[columnName] = "Age must be greater than 30";
                                        }
                                        break;
                                    case "value != null && value.Trim() != ''":
                                        if (string.IsNullOrEmpty(cellValue?.Trim()))
                                        {
                                            rowValidation[columnName] = "Value must not be empty";
                                        }
                                        break;
                                    // Add more validation checks as needed
                                    default:
                                        break;
                                }
                            }
                        }
                        data.Add(rowData);
                        validationResults.Add(rowValidation);
                    });
                    stopwatch.Stop();
                    _logger.LogInformation($"Processed rows in {stopwatch.ElapsedMilliseconds} ms.");
                }
            }

            return Ok(new { columns = excelColumns.Select(col => new { Header = col }), data, validationResults });
        }
    }
}


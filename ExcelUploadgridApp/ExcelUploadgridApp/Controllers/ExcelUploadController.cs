//using ExcelUploadgridApp.Data;
//using ExcelUploadgridApp.Model;
//using Microsoft.AspNetCore.Http;
//using Microsoft.AspNetCore.Mvc;
//using Microsoft.EntityFrameworkCore;
//using OfficeOpenXml;
//using System.Text.Json;

//namespace ExcelUploadgridApp.Controllers
//{
//    [Route("api/[controller]")]
//    [ApiController]
//    public class ExcelUploadController : ControllerBase
//    {
//        private readonly AppDbContext _context;

//        public ExcelUploadController(AppDbContext context)
//        {
//            _context = context;
//        }



//        [HttpGet("datasources")]
//        public async Task<IActionResult> GetDatasources()
//        {
//            var datasources = await _context.ParentDetails
//                                            .Where(d => d.IsActive)
//                                            .Select(d => new { d.Id, d.DatasourceName })
//                                            .ToListAsync();
//            return Ok(datasources);
//        }

//        [HttpGet("columns/{datasourceId}")]
//        public async Task<IActionResult> GetColumns(int datasourceId)
//        {
//            var columns = await _context.ColumnDetails
//                                        .Where(c => c.ParentId == datasourceId)
//                                        .Select(c => new { c.ColumnName, c.DataType })
//                                        .ToListAsync();
//            return Ok(columns);
//        }


//        //basic upload code for excel
//        //[HttpPost("upload")]
//        //public IActionResult Upload(IFormFile file)
//        //{
//        //    using var package = new ExcelPackage(file.OpenReadStream());
//        //    var worksheet = package.Workbook.Worksheets[0];
//        //    var rowCount = worksheet.Dimension.Rows;
//        //    var colCount = worksheet.Dimension.Columns;

//        //    var columns = new List<object>();
//        //    for (var col = 1; col <= colCount; col++)
//        //    {
//        //        columns.Add(new { Header = worksheet.Cells[1, col].Text });
//        //    }

//        //    var data = new List<object>();
//        //    for (var row = 2; row <= rowCount; row++)
//        //    {
//        //        var rowData = new Dictionary<string, object>();
//        //        for (var col = 1; col <= colCount; col++)
//        //        {
//        //            rowData[worksheet.Cells[1, col].Text] = worksheet.Cells[row, col].Text;
//        //        }
//        //        data.Add(rowData);
//        //    }

//        //    return Ok(new { columns, data });
//        //}

//        //[HttpPost("upload")]
//        //public async Task<IActionResult> Upload(IFormFile file, [FromQuery] int datasourceId)
//        //{
//        //    var columns = await _context.ColumnDetails
//        //                                .Where(c => c.ParentId == datasourceId)
//        //                                .Select(c => c.UserFriendlyName)
//        //                                .ToListAsync();

//        //    using var package = new ExcelPackage(file.OpenReadStream());
//        //    var worksheet = package.Workbook.Worksheets[0];
//        //    var rowCount = worksheet.Dimension.Rows;
//        //    var colCount = worksheet.Dimension.Columns;

//        //    var excelColumns = new List<string>();
//        //    for (var col = 1; col <= colCount; col++)
//        //    {
//        //        excelColumns.Add(worksheet.Cells[1, col].Text);
//        //    }

//        //    var missingColumns = columns.Except(excelColumns).ToList();
//        //    if (missingColumns.Any())
//        //    {
//        //        return BadRequest(new { message = "Excel file is missing required columns", missingColumns });
//        //    }

//        //    var data = new List<Dictionary<string, object>>();
//        //    for (var row = 2; row <= rowCount; row++)
//        //    {
//        //        var rowData = new Dictionary<string, object>();
//        //        for (var col = 1; col <= colCount; col++)
//        //        {
//        //            rowData[worksheet.Cells[1, col].Text] = worksheet.Cells[row, col].Text;
//        //        }
//        //        data.Add(rowData);
//        //    }

//        //    return Ok(new { columns = excelColumns.Select(col => new { Header = col }), data });
//        //}

//        [HttpPost("upload")]
//        public async Task<IActionResult> Upload(IFormFile file, [FromQuery] int datasourceId)
//        {
//            var columns = await _context.ColumnDetails
//                                        .Where(c => c.ParentId == datasourceId)
//                                        .Select(c => new {
//                                            c.UserFriendlyName,
//                                            c.ConstraintExpression,
//                                            c.ColumnName
//                                        })
//                                        .ToListAsync();

//            using var package = new ExcelPackage(file.OpenReadStream());
//            var worksheet = package.Workbook.Worksheets[0];
//            var rowCount = worksheet.Dimension.Rows;
//            var colCount = worksheet.Dimension.Columns;

//            var excelColumns = new List<string>();
//            for (var col = 1; col <= colCount; col++)
//            {
//                excelColumns.Add(worksheet.Cells[1, col].Text);
//            }

//            var missingColumns = columns.Select(c => c.UserFriendlyName).Except(excelColumns).ToList();
//            if (missingColumns.Any())
//            {
//                return BadRequest(new { message = "Excel file is missing required columns", missingColumns });
//            }

//            var data = new List<Dictionary<string, object>>();
//            var validationResults = new List<Dictionary<string, string>>();

//            for (var row = 2; row <= rowCount; row++)
//            {
//                var rowData = new Dictionary<string, object>();
//                var rowValidation = new Dictionary<string, string>();
//                for (var col = 1; col <= colCount; col++)
//                {
//                    var cellValue = worksheet.Cells[row, col].Text;
//                    var columnName = worksheet.Cells[1, col].Text;
//                    rowData[columnName] = cellValue;

//                    var column = columns.FirstOrDefault(c => c.UserFriendlyName == columnName);
//                    if (column != null && !string.IsNullOrEmpty(column.ConstraintExpression))
//                    {
//                        // Example validation checks based on column.ConstraintExpression
//                        switch (column.ConstraintExpression)
//                        {
//                            case "value > 100":
//                                if (int.TryParse(cellValue, out var intValue) && intValue <= 100)
//                                {
//                                    rowValidation[columnName] = "Value must be greater than 100";
//                                }
//                                break;
//                            case "value <= DateTime.Now":
//                                if (DateTime.TryParse(cellValue, out var dateValue) && dateValue > DateTime.Now)
//                                {
//                                    rowValidation[columnName] = "Date must be in the past";
//                                }
//                                break;
//                            case "value == 0 || value == 1":
//                                if (int.TryParse(cellValue, out var bitValue) && (bitValue != 0 && bitValue != 1))
//                                {
//                                    rowValidation[columnName] = "Value must be 0 or 1";
//                                }
//                                break;
//                            case "value > 30":
//                                if (int.TryParse(cellValue, out var ageValue) && ageValue <= 30)
//                                {
//                                    rowValidation[columnName] = "Age must be greater than 30";
//                                }
//                                break;
//                            case "value != null && value.Trim() != ''":
//                                if (string.IsNullOrEmpty(cellValue?.Trim()))
//                                {
//                                    rowValidation[columnName] = "Value must not be empty";
//                                }
//                                break;
//                            // Add more validation checks as needed
//                            default:
//                                break;
//                        }
//                    }
//                }
//                data.Add(rowData);
//                validationResults.Add(rowValidation);
//            }

//            return Ok(new { columns = excelColumns.Select(col => new { Header = col }), data, validationResults });
//        }

//        //[HttpPost("save")]
//        //public async Task<IActionResult> SaveValues(int dataSourceId, [FromBody] List<Dictionary<string, object>> editedValues)
//        //{
//        //    var columnDetails = await _context.ColumnDetails
//        //                                .Where(c => c.ParentId == dataSourceId)
//        //                                .Select(c => new {
//        //                                    c.UserFriendlyName,
//        //                                    c.ConstraintExpression,
//        //                                    c.ColumnName,
//        //                                    c.Id
//        //                                })
//        //                                .ToListAsync();
//        //    try
//        //    {
//        //        int lastRowId = await _context.ValueDetails.AnyAsync() ? await _context.ValueDetails.MaxAsync(v => v.RowId) : 0;

//        //        foreach (var editedRow in editedValues)
//        //        {
//        //            int rowId = editedRow.ContainsKey("__rowId") ? (int)editedRow["__rowId"] : ++lastRowId;

//        //            foreach (var editedCell in editedRow)
//        //            {
//        //                if (editedCell.Key == "__rowId") continue;  // Skip the row ID key

//        //                var columnName = editedCell.Key.ToLower();
//        //                var cellValue = editedCell.Value;

//        //                var column = columnDetails.FirstOrDefault(c => c.ColumnName.ToLower() == columnName);

//        //                if (column != null)
//        //                {
//        //                    // Find the corresponding row in valuedetails table using the row ID
//        //                    var valueDetail = await _context.ValueDetails.FirstOrDefaultAsync(v => v.RowId == rowId && v.ColumnId == column.Id);

//        //                    if (valueDetail != null)
//        //                    {
//        //                        // Update the value in valuedetails table
//        //                        valueDetail.Value = cellValue.ToString();
//        //                        _context.ValueDetails.Update(valueDetail);
//        //                    }
//        //                    else
//        //                    {
//        //                        // Insert a new row in valuedetails table
//        //                        _context.ValueDetails.Add(new ValueDetail
//        //                        {
//        //                            ColumnId = column.Id,
//        //                            RowId = rowId,
//        //                            Value = cellValue.ToString()
//        //                        });
//        //                    }
//        //                }
//        //            }

//        //            if (!editedRow.ContainsKey("__rowId"))
//        //            {
//        //                editedRow["__rowId"] = rowId;
//        //            }
//        //        }

//        //        // Save changes to the database
//        //        await _context.SaveChangesAsync();

//        //        return Ok(new { message = "Values saved successfully" });
//        //    }
//        //    catch (Exception ex)
//        //    {
//        //        return StatusCode(StatusCodes.Status500InternalServerError, new { message = "Failed to save values", error = ex.Message });
//        //    }
//        //}


//        [HttpPost("save")]
//public async Task<IActionResult> SaveValues(int dataSourceId, [FromBody] List<Dictionary<string, object>> editedValues)
//{
//    var columnDetails = await _context.ColumnDetails
//                                .Where(c => c.ParentId == dataSourceId)
//                                .Select(c => new {
//                                    c.UserFriendlyName,
//                                    c.ConstraintExpression,
//                                    c.ColumnName,
//                                    c.Id
//                                })
//                                .ToListAsync();
//    try
//    {
//        foreach (var editedRow in editedValues)
//        {
//                    int rowId = editedRow.ContainsKey("rowId") ? ((JsonElement)editedRow["rowId"]).GetInt32() : 0;

//                    foreach (var editedCell in editedRow)
//            {
//                if (editedCell.Key == "rowId") continue;

//                var columnName = editedCell.Key.ToLower();
//                var cellValue = editedCell.Value;

//                var column = columnDetails.FirstOrDefault(c => c.ColumnName.ToLower() == columnName);

//                if (column != null)
//                {
//                    // Find the corresponding row in valuedetails table using the row ID
//                    var valueDetail = await _context.ValueDetails.FirstOrDefaultAsync(v => v.RowId == rowId && v.ColumnId == column.Id);

//                    if (valueDetail != null)
//                    {
//                        // Update the value in valuedetails table
//                        valueDetail.Value = cellValue.ToString();
//                        _context.ValueDetails.Update(valueDetail);
//                    }
//                    else
//                    {
//                        // Insert a new row in valuedetails table
//                        _context.ValueDetails.Add(new ValueDetail
//                        {
//                            ColumnId = column.Id,
//                            RowId = rowId,
//                            Value = cellValue.ToString()
//                        });
//                    }
//                }
//            }
//        }

//        // Save changes to the database
//        await _context.SaveChangesAsync();

//        return Ok(new { message = "Values saved successfully" });
//    }
//    catch (Exception ex)
//    {
//        return StatusCode(StatusCodes.Status500InternalServerError, new { message = "Failed to save values", error = ex.Message });
//    }
//}


//    }
//}
//Siddhesh Ramesh Ahire
using ExcelUploadgridApp.Data;
using ExcelUploadgridApp.Model;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.Json;
using System.Linq;

namespace ExcelUploadgridApp.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelUploadController : ControllerBase
    {
        private readonly AppDbContext _context;

        public ExcelUploadController(AppDbContext context)
        {
            _context = context;
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
            var columns = await _context.ColumnDetails
                                        .Where(c => c.ParentId == datasourceId)
                                        .Select(c => new {
                                            c.UserFriendlyName,
                                            c.ConstraintExpression,
                                            c.ColumnName
                                        })
                                        .ToListAsync();

            var excelColumns = new List<string>();
            var data = new List<Dictionary<string, object>>();
            var validationResults = new List<Dictionary<string, string>>();

            using (var stream = file.OpenReadStream())
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Open(stream, false))
                {
                    var workbookPart = spreadsheetDocument.WorkbookPart;
                    var sheet = workbookPart.Workbook.Sheets.Cast<Sheet>().FirstOrDefault();
                    var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                    var sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    var rows = sheetData.Elements<Row>().ToList();
                    if (rows.Count == 0)
                    {
                        return BadRequest(new { message = "The Excel file is empty" });
                    }

                    var headerRow = rows.First();
                    var cellValues = headerRow.Elements<Cell>().Select(c => GetCellValue(c, workbookPart)).ToList();
                    excelColumns.AddRange(cellValues);

                    var missingColumns = columns.Select(c => c.UserFriendlyName).Except(excelColumns).ToList();
                    if (missingColumns.Any())
                    {
                        return BadRequest(new { message = "Excel file is missing required columns", missingColumns });
                    }

                    foreach (var row in rows.Skip(1))
                    {
                        var rowData = new Dictionary<string, object>();
                        var rowValidation = new Dictionary<string, string>();

                        for (var colIndex = 0; colIndex < row.Elements<Cell>().Count(); colIndex++)
                        {
                            var cell = row.Elements<Cell>().ElementAt(colIndex);
                            var columnName = excelColumns[colIndex];
                            var cellValue = GetCellValue(cell, workbookPart);

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
                    }
                }
            }

            return Ok(new { columns = excelColumns.Select(col => new { Header = col }), data, validationResults });
        }

        private string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null || cell.CellValue == null)
            {
                return string.Empty;
            }

            var value = cell.CellValue.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (stringTable != null)
                {
                    return stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }

            return value;
        }

        [HttpPost("save")]
        public async Task<IActionResult> SaveValues(int dataSourceId, [FromBody] List<Dictionary<string, object>> editedValues)
        {
            var columnDetails = await _context.ColumnDetails
                                        .Where(c => c.ParentId == dataSourceId)
                                        .Select(c => new {
                                            c.UserFriendlyName,
                                            c.ConstraintExpression,
                                            c.ColumnName,
                                            c.Id
                                        })
                                        .ToListAsync();
            try
            {
                foreach (var editedRow in editedValues)
                {
                    int rowId = editedRow.ContainsKey("rowId") ? ((JsonElement)editedRow["rowId"]).GetInt32() : 0;

                    foreach (var editedCell in editedRow)
                    {
                        if (editedCell.Key == "rowId") continue;

                        var columnName = editedCell.Key.ToLower();
                        var cellValue = editedCell.Value;

                        var column = columnDetails.FirstOrDefault(c => c.ColumnName.ToLower() == columnName);

                        if (column != null)
                        {
                            // Find the corresponding row in valuedetails table using the row ID
                            var valueDetail = await _context.ValueDetails.FirstOrDefaultAsync(v => v.RowId == rowId && v.ColumnId == column.Id);

                            if (valueDetail != null)
                            {
                                // Update the value in valuedetails table
                                valueDetail.Value = cellValue.ToString();
                                _context.ValueDetails.Update(valueDetail);
                            }
                            else
                            {
                                // Insert a new row in valuedetails table
                                _context.ValueDetails.Add(new ValueDetail
                                {
                                    ColumnId = column.Id,
                                    RowId = rowId,
                                    Value = cellValue.ToString()
                                });
                            }
                        }
                    }
                }

                // Save changes to the database
                await _context.SaveChangesAsync();

                return Ok(new { message = "Values saved successfully" });
            }
            catch (Exception ex)
            {
                return StatusCode(StatusCodes.Status500InternalServerError, new { message = "Failed to save values", error = ex.Message });
            }
        }
    }
}


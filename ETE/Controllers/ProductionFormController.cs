using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ETE.Dtos;
using ETE.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using System.Diagnostics;
using System.Reflection.PortableExecutable;
using System.Text.Json;

namespace ETE.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ProductionFormController : ControllerBase
    {
        private readonly AppDbContext _context;

        public ProductionFormController(AppDbContext context)
        {
            _context = context;
        }

        [HttpGet]
        [Route("GetListTimes")]
        public async Task<IActionResult> GetListTimes()
        {
            var list = await _context.Hours.AsNoTracking().ToListAsync();           

            return Ok(list);
        }

        [HttpGet]
        [Route("GetLines")]
        public async Task<IActionResult> GetLines()
        {
            var lines = await _context.Lines.AsNoTracking().ToListAsync();

            return Ok(lines);
        }

        [HttpGet]
        [Route("GetCodes")]
        public async Task<IActionResult> GetCodes()
        {
            var codes = await _context.Codes.AsNoTracking().ToListAsync();

            return Ok(codes);
        }

        [HttpGet]
        [Route("GetWorkShifts")]
        public async Task<IActionResult> GetWorkShifts()
        {
            var workShifts = await _context.WorkShift.AsNoTracking().ToListAsync();

            return Ok(workShifts);
        }

        [HttpGet]
        [Route("GetMachines")]
        public async Task<IActionResult> GetMachines()
        {
            var machines = await _context.Machine.AsNoTracking().ToListAsync();
            return Ok(machines);
        }

        [HttpGet]
        [Route("GetProcessesByLine/{lineId}")]
        public async Task<IActionResult> GetProcessesBy(int lineId)
        {
            var processes = await _context.Process
                                .Where(p => p.LineId == lineId)
                                .AsNoTracking()
                                .ToListAsync();

            return Ok(processes);
        }

        [HttpGet]
        [Route("GetMachinesByProcess/{processId}")]
        public async Task<IActionResult> GetMachinesByProcess(int processId)
        {
            var machines = await _context.Machine
                            .Where(m => m.ProcessId == processId)
                            .AsNoTracking()
                            .ToListAsync();

            return Ok(machines);
        }

        [HttpGet]
        [Route("GetReasonByCodes/{codesId}")]
        public async Task<IActionResult> GetReasonByCodes(int codesId)
        {
            var reasons = await _context.Reason
                                .Where(r => r.CodeId == codesId)
                                .AsNoTracking()
                                .ToListAsync();

            return Ok(reasons);
        }

        [HttpGet]
        [Route("GetMachineByLine/{lineId}")]
        public async Task<IActionResult> GetMachineByLine(int lineId)
        {
            var machines = await _context.Machine
                                .Where(m => m.Process.LineId == lineId)
                                .AsNoTracking()
                                .ToListAsync();

            return Ok(machines);
        }

        [HttpGet]
        [Route("GetQualityData")]
        public async Task<IActionResult> GetQualityData(
            [FromQuery] int? lineId,
            [FromQuery] int? shiftId,
            [FromQuery] int? machineId,
            [FromQuery] DateTime? startDate = null,
            [FromQuery] DateTime? endDate = null)
        {
            try
            {
                if (!startDate.HasValue) startDate = DateTime.Now.AddDays(-30);

                if (!endDate.HasValue) endDate = DateTime.Now;

                var hourIdsQuery = _context.Hours
                    .Where(h => h.Date >= startDate && h.Date <= endDate)
                    .AsNoTracking()
                    .AsQueryable();

                var filterProduction = _context.Production.AsQueryable();
                var filteredWorkShift = _context.WorkShiftHours.AsQueryable();
               

                if (shiftId.HasValue)
                {
                    filteredWorkShift = filteredWorkShift.Where(wsh => wsh.WorkShiftId == shiftId.Value);
                }

                var hourIds = await hourIdsQuery
                    .Select(h => h.Id)
                    .ToListAsync();

                Console.WriteLine($"IDs de horas encontrados: {hourIds.Count}");

                var productionQuery = _context.Production
                    .Where(p => hourIds.Contains(p.HourId))
                    .AsQueryable();

                if (lineId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => p.LinesId == lineId.Value);
                }

                if (machineId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => p.MachineId == machineId.Value);
                }

                var totals = await productionQuery
                    .GroupBy(p => 1)
                    .Select(g => new
                    {
                        TotalPieces = g.Sum(p => p.PieceQuantity ?? 0),
                        TotalScrap = g.Sum(p => p.Scrap ?? 0)
                    })
                    .FirstOrDefaultAsync();

                if (totals == null || (totals.TotalPieces == 0 && totals.TotalScrap == 0))
                {                    
                    if (!lineId.HasValue && !machineId.HasValue && !shiftId.HasValue)
                    {
                        productionQuery = _context.Production.AsQueryable();
                        totals = await productionQuery
                            .GroupBy(p => 1)
                            .Select(g => new
                            {
                                TotalPieces = g.Sum(p => p.PieceQuantity ?? 0),
                                TotalScrap = g.Sum(p => p.Scrap ?? 0)
                            })
                            .FirstOrDefaultAsync();
                    }
                }

                var result = new
                {
                    GoodPieces = totals?.TotalPieces - totals?.TotalScrap ?? 0,
                    Scrap = totals?.TotalScrap ?? 0,
                    TotalPieces = totals?.TotalPieces ?? 0,
                    QualityPercentage = totals == null || totals.TotalPieces == 0 ? 0 :
                        Math.Round(((totals.TotalPieces - totals.TotalScrap) * 100.0) / totals.TotalPieces, 2)
                };

                Console.WriteLine($"Resultados: {JsonSerializer.Serialize(result)}");

                return Ok(result);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.ToString()}");
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetEfficiencyData")]
        public async Task<IActionResult> GetEfficiencyData(
            [FromQuery] int? lineId,
            [FromQuery] int? shiftId,
            [FromQuery] int? machineId,
            [FromQuery] DateTime? startDate = null,
            [FromQuery] DateTime? endDate = null)
        {
            try
            {
                if (!startDate.HasValue) startDate = DateTime.Now.AddDays(-30);

                if (!endDate.HasValue) endDate = DateTime.Now;

                var productionQuery = _context.Production
                            .Where(p => p.RegistrationDate >= startDate && p.RegistrationDate <= endDate);

                if (lineId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => p.LinesId == lineId.Value);
                }

                if (machineId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => p.MachineId == machineId.Value);
                }

                if (shiftId.HasValue)
                {
                    productionQuery = productionQuery.Where(p => _context.WorkShiftHours
                        .Any(wsh => wsh.HourId == p.HourId && wsh.WorkShiftId == shiftId));
                }


                var masterQuery = _context.MasterEngineering
                        .Where(me => me.PzHr != null && me.PzHr > 0)
                        .GroupBy(me => new { me.ParentPartNumber, me.Line })
                        .Select(g => new
                        {
                            ParentPartNumber = g.Key.ParentPartNumber,
                            Line = g.Key.Line,
                            AvergaePzHr = g.Average(me => me.PzHr)
                        });

                var efficiencyQuery = productionQuery
                    .Join(masterQuery,
                        p => new { PartNumber = p.PartNumber, LineId = p.LinesId },
                        m => new { PartNumber = m.ParentPartNumber, LineId = m.Line },
                        (p, m) => new
                        {
                            p.Id,
                            p.PieceQuantity,
                            m.AvergaePzHr,
                            p.RegistrationDate,
                            p.HourId,
                            p.PartNumber,
                            p.LinesId,
                            m.ParentPartNumber,
                            m.Line
                        })
                    .Join(_context.Hours,
                        x => x.HourId,
                        h => h.Id,
                        (x, h) => new
                        {
                            x.Id,
                            x.PieceQuantity,
                            x.AvergaePzHr,
                            x.RegistrationDate,
                            HourDate = h.Date,
                            x.PartNumber,
                            x.LinesId,
                            x.ParentPartNumber,
                            x.Line
                    });

                var groupedData = await efficiencyQuery
                        .GroupBy(x => new
                        {
                            Date = x.HourDate.HasValue ? x.HourDate.Value.Date : x.RegistrationDate.Value.Date,
                        })
                        .Select(g => new
                        {
                            Date = g.Key.Date,
                            TotalProduced = g.Sum(x => x.PieceQuantity) ?? 0,
                            TotalExpected = g.Sum(x => x.AvergaePzHr) ?? 1,
                            Efficiency = (g.Sum(x => x.PieceQuantity) ?? 0) / (double)(g.Sum(x => x.AvergaePzHr) ?? 1),
                            Details = g.Select(x => new
                            {
                                x.PartNumber,
                                x.LinesId,
                                x.ParentPartNumber,
                                x.Line,
                                x.PieceQuantity,
                                x.AvergaePzHr
                            }).ToList()
                        })
                        .OrderBy(x => x.Date)
                        .ToListAsync();

                return Ok(new
                {
                    Summary = groupedData,
                    Message = "Nota: Se usa el promedio de pzHr cuando hay multiples valores para la misma combinación partNumber/Linea"
                });
            }
            catch(Exception ex)
            {
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetAvailabilityData")]
        public async Task<IActionResult> GetAvailabilityDataNoDateFilter(
            [FromQuery] int? lineId,
            [FromQuery] int? shiftId,
            [FromQuery] int? machineId,
            [FromQuery] DateTime? startDate = null,
            [FromQuery] DateTime? endDate = null)
        {
            try
            {
                var hourQuery = _context.Hours.AsQueryable();

                var productionQuery = _context.Production.AsQueryable();

                if (!startDate.HasValue) startDate = DateTime.Now.AddDays(-30);

                if (!endDate.HasValue) endDate = DateTime.Now;

                if (shiftId.HasValue)
                {
                    hourQuery = hourQuery
                        .Where(h => _context.WorkShiftHours
                            .Any(wsh => wsh.HourId == h.Id && wsh.WorkShiftId == shiftId.Value));
                }               

                var filteredHourIds = await hourQuery.Select(h => h.Id).ToListAsync();

                if (!filteredHourIds.Any())
                {
                    return Ok(new
                    {
                        totalTime = 0,
                        deadTime = 0,
                        availableTime = 0,
                        percentage = 0
                    });
                }

                int totalMinutes = filteredHourIds.Count * 60;

                var deadTimesQuery = _context.Production
                    .Where(p => filteredHourIds.Contains(p.HourId))
                    .Join(_context.DeadTimes,
                        p => p.DeadTimesId,
                        dt => dt.Id,
                        (p, dt) => new { dt.Minutes, p.LinesId, p.MachineId });

                if (lineId.HasValue)
                {
                    deadTimesQuery = deadTimesQuery.Where(x => x.LinesId == lineId.Value);
                }

                if (machineId.HasValue)
                {
                    deadTimesQuery = deadTimesQuery.Where(x => x.MachineId == machineId.Value);
                }

                int totalDeadTime = await deadTimesQuery.SumAsync(x => x.Minutes ?? 0);
                int availableTime = totalMinutes - totalDeadTime;
                double availabilityPercentage = totalMinutes > 0
                    ? Math.Round((double)availableTime / totalMinutes * 100, 2)
                    : 0;

                return Ok(new
                {
                    totalTime = totalMinutes,
                    deadTime = totalDeadTime,
                    availableTime = availableTime,
                    percentage = availabilityPercentage
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Error al calcular la disponibilidad: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetDeadTimeByReasonLast6Days")]
        public async Task<IActionResult> GetDeadTimeByReasonLast6Days(
            [FromQuery] int? lineId = null,
            [FromQuery] int? shiftId = null,
            [FromQuery] int? machineId = null,
            [FromQuery] DateTime? startDate = null,
            [FromQuery] DateTime? endDate = null)
        {
            try
            {
                if (!startDate.HasValue) startDate = DateTime.Now.AddDays(-30);

                if (!endDate.HasValue) endDate = DateTime.Now;

                var deadTimeQuery = _context.DeadTimes.AsQueryable();
                var productionQuery = _context.Production.AsQueryable();                

                if (lineId.HasValue)
                {
                    deadTimeQuery = deadTimeQuery.Where(dt =>
                        _context.Production.Any(p => p.DeadTimesId == dt.Id && p.LinesId == lineId));

                    productionQuery = productionQuery.Where(p => p.LinesId == lineId);
                }

                if (machineId.HasValue)
                {
                    deadTimeQuery = deadTimeQuery.Where(dt =>
                        _context.Production.Any(p => p.DeadTimesId == dt.Id && p.MachineId == machineId));

                    productionQuery = productionQuery.Where(p => p.MachineId == machineId);
                }

                if (shiftId.HasValue)
                {
                    deadTimeQuery = deadTimeQuery.Where(dt =>
                        _context.Production.Any(p => p.DeadTimesId == dt.Id &&
                            _context.WorkShiftHours.Any(wsh =>
                                wsh.HourId == p.HourId && wsh.WorkShiftId == shiftId)));

                    productionQuery = productionQuery.Where(p =>
                        _context.WorkShiftHours.Any(wsh =>
                            wsh.HourId == p.HourId && wsh.WorkShiftId == shiftId));
                }

                var result = await _context.Production
                    .Where(p => p.Hour.Date >= startDate &&
                               p.Hour.Date <= endDate &&
                               p.DeadTimesId != null &&
                               (lineId == null || p.LinesId == lineId))
                    .GroupBy(p => p.DeadTimes.Reason.Name)
                    .Select(g => new {
                        Reason = g.Key ?? "Sin razón",
                        TotalMinutes = g.Sum(p => p.DeadTimes.Minutes ?? 0)
                    })
                    .OrderByDescending(x => x.TotalMinutes)
                    .ToListAsync();
               
                if (!result.Any())
                {
                    result = await _context.DeadTimes
                        .GroupBy(dt => dt.Reason.Name)
                        .Select(g => new {
                            Reason = g.Key ?? "Sin razón",
                            TotalMinutes = g.Sum(dt => dt.Minutes ?? 0)
                        })
                        .OrderByDescending(x => x.TotalMinutes)
                        .ToListAsync();
                }

                var totalMinutes = result.Sum(x => x.TotalMinutes);
                var averageMinutes = totalMinutes / 6.0;

                return Ok(new
                {
                    labels = result.Select(x => x.Reason).ToArray(),
                    data = result.Select(x => x.TotalMinutes).ToArray(),
                    totalMinutes,
                    averageMinutes
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetKeyMetrics")]
        public async Task<IActionResult> GetKeyMetrics(
            [FromQuery] int? lineId = null,
            [FromQuery] int? shiftId = null,
            [FromQuery] int? machineId = null,
            [FromQuery] DateTime? startDate = null,
            [FromQuery] DateTime? endDate = null)
        {
            try
            {

                var deadTimeQuery = _context.DeadTimes.AsQueryable();
                var productionQuery = _context.Production.AsQueryable();
                
                if (!startDate.HasValue) startDate = DateTime.Now.AddDays(-30);

                if (!endDate.HasValue) endDate = DateTime.Now;

                if (lineId.HasValue)
                {
                    deadTimeQuery = deadTimeQuery.Where(dt =>
                        _context.Production.Any(p => p.DeadTimesId == dt.Id && p.LinesId == lineId));

                    productionQuery = productionQuery.Where(p => p.LinesId == lineId);
                }

                if (machineId.HasValue)
                {
                    deadTimeQuery = deadTimeQuery.Where(dt =>
                        _context.Production.Any(p => p.DeadTimesId == dt.Id && p.MachineId == machineId));

                    productionQuery = productionQuery.Where(p => p.MachineId == machineId);
                }

                if (shiftId.HasValue)
                {
                    deadTimeQuery = deadTimeQuery.Where(dt =>
                        _context.Production.Any(p => p.DeadTimesId == dt.Id &&
                            _context.WorkShiftHours.Any(wsh =>
                                wsh.HourId == p.HourId && wsh.WorkShiftId == shiftId)));

                    productionQuery = productionQuery.Where(p =>
                        _context.WorkShiftHours.Any(wsh =>
                            wsh.HourId == p.HourId && wsh.WorkShiftId == shiftId));
                }

                var totalDeadTime = await deadTimeQuery.SumAsync(dt => dt.Minutes ?? 0);
                var totalScrap = await productionQuery.SumAsync(p => p.Scrap ?? 0);

                var avgDeadTimeQuery = _context.DeadTimes.AsQueryable();
                var avgScrapQuery = _context.Production.AsQueryable();

                if (startDate.HasValue || endDate.HasValue)
                {
                    var startDateValue = startDate ?? DateTime.MinValue;
                    var endDateValue = endDate ?? DateTime.MaxValue;
                    endDateValue = endDateValue.Date.AddDays(1).AddTicks(-1);

                    avgDeadTimeQuery = avgDeadTimeQuery.Where(dt =>
                        dt.RegistrationDate >= startDateValue &&
                        dt.RegistrationDate <= endDateValue);

                    avgScrapQuery = avgScrapQuery.Where(p =>
                        p.RegistrationDate >= startDateValue &&
                        p.RegistrationDate <= endDateValue);
                }

                var avgDeadTime = await avgDeadTimeQuery.AverageAsync(dt => dt.Minutes ?? 0);
                var avgScrap = await avgScrapQuery.AverageAsync(p => p.Scrap ?? 0);

                return Ok(new
                {
                    DeadTime = totalDeadTime,
                    DeadTimeVsAvg = totalDeadTime - avgDeadTime,
                    Scrap = totalScrap,
                    ScrapVsAvg = totalScrap - avgScrap,
                    Metadata = new
                    {
                        StartDate = startDate,
                        EndDate = endDate,
                        LineFilter = lineId,
                        MachineFilter = machineId,
                        ShiftFilter = shiftId
                    }
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"Internal server error: {ex.Message}");
            }
        }

        [HttpGet]
        [Route("GetListProduction")]
        public async Task<IActionResult> GetLisProduction()
        {
            var list = await _context.Production
                .Select(p => new
                {
                    p.Id,
                    p.RegistrationDate,
                    Line = p.Lines.Name,
                    Machine = p.Process.Machine.FirstOrDefault()!.Name,
                    Hour = p.Hour.Time,
                    p.PieceQuantity,
                    p.DeadTimes.Minutes,
                    p.Scrap
                })
                .AsNoTracking()
                .ToListAsync();

            return Ok(list);
        }

        [HttpPost]
        [Route("DownloadProduction")]
        public async Task<IActionResult> ExportToExcel()
        {
            var production = await _context.Production.Select(p => new
            {
                p.Id,
                p.RegistrationDate,
                Line = p.Lines.Name,
                Machine = p.Process.Machine.FirstOrDefault()!.Name,
                Hour = p.Hour.Time,
                p.PieceQuantity,
                p.DeadTimes.Minutes,
                p.Scrap
            })
            .AsNoTracking()
            .ToListAsync();

            using (var workBook = new XLWorkbook())
            {
                var workSheet = workBook.Worksheets.Add("Production");

                workSheet.Cell(1, 1).Value = "Fecha";
                workSheet.Cell(1, 2).Value = "Línea";
                workSheet.Cell(1, 3).Value = "Máquina";
                workSheet.Cell(1, 4).Value = "Hora";
                workSheet.Cell(1, 5).Value = "Cantidad de Piezas";
                workSheet.Cell(1, 6).Value = "Tiempo Muerto (Minutos)";
                workSheet.Cell(1, 7).Value = "Scrap";

                for(int i = 0; i <  production.Count; i++)
                {
                    workSheet.Cell(i + 2, 1).Value = production[i].RegistrationDate;
                    workSheet.Cell(i + 2, 2).Value = production[i].Line;
                    workSheet.Cell(i + 2, 3).Value = production[i].Machine;
                    workSheet.Cell(i + 2, 4).Value = production[i].Hour;
                    workSheet.Cell(i + 2, 5).Value = production[i].PieceQuantity;
                    workSheet.Cell(i + 2, 6).Value = production[i].Minutes;
                    workSheet.Cell(i + 2, 7).Value = production[i].Scrap;
                }

                var stream = new MemoryStream();
                workBook.SaveAs(stream);
                stream.Position = 0;

                var fileName = $"DatosExportados_{DateTime.Now:ddMMyyyy}.xlsx";

                return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
            }
        }

        [HttpPost]
        [Route("RegisterProduction")]
        public async Task<IActionResult> RegisterProduction([FromBody] ProductionRegisterDto productionDto)
        {
            try
            {
                if (!ModelState.IsValid)
                {
                    return BadRequest(ModelState);
                }

                if(productionDto.DeadTimes != null && productionDto.DeadTimes.Any())
                {
                    foreach(var dt in productionDto.DeadTimes)
                    {
                        var codeExists = await _context.Codes.AnyAsync(c => c.Id == dt.CodeId);
                        var reasonExists = await _context.Reason.AnyAsync(r => r.Id == dt.ReasonId);

                        if(!codeExists || !reasonExists)
                        {
                            return BadRequest("Código o razón de tiempo muerto no válido");
                        }
                    }
                }

                List<DeadTimes> createdDeadTimes = new List<DeadTimes>();
                if(productionDto.DeadTimes != null && productionDto.DeadTimes.Any())
                {
                    foreach(var deadTimeDto in productionDto.DeadTimes)
                    {
                        var deadTime = new DeadTimes
                        {
                            Minutes = deadTimeDto.Minutes,
                            CodeId = deadTimeDto.CodeId,
                            ReasonId = deadTimeDto.ReasonId,
                        };

                        _context.DeadTimes.Add(deadTime);
                        createdDeadTimes.Add(deadTime);
                    }
                    await _context.SaveChangesAsync();
                }

                var production = new Production
                {
                    PartNumber = productionDto.PartNumber,
                    PieceQuantity = productionDto.PieceQuantity,
                    HourId = productionDto.HourId,
                    LinesId = productionDto.LinesId,
                    ProcessId = productionDto.ProcessId,
                    MachineId = productionDto.MachineId,
                    Scrap = productionDto.Scrap,
                };

                if (createdDeadTimes.Any())
                {
                    foreach(var dt in createdDeadTimes)
                    {
                        production.DeadTimesId = dt.Id;
                    }
                }

                _context.Production.Add(production);
                await _context.SaveChangesAsync();

                return Ok(new
                {
                    Succes = true,
                    ProductionId = production.Id,
                    DeadTimesIds = createdDeadTimes.Select(dt => dt.Id).ToList()
                });
            }
            catch(Exception ex)
            {
                return StatusCode(500, $"Error interno del servidor: {ex.Message}");
            }
        }
    }
}
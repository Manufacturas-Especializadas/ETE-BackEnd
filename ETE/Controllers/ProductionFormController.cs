using ETE.Dtos;
using ETE.Models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

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
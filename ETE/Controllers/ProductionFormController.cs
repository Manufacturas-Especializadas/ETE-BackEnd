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
    }
}
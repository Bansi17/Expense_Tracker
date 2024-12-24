using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Expense_Tracker.Data;
using Expense_Tracker.Models;
using Expense_Tracker.DTOs;

namespace Expense_Tracker.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class LoginController : ControllerBase
    {
        private readonly ApplicationDbContext _context;

        public LoginController(ApplicationDbContext context)
        {
            _context = context;
        }

        // POST: api/Login
        [HttpPost]
        public async Task<IActionResult> Login([FromBody] UserLoginDto login)
        {
            if (login == null || string.IsNullOrEmpty(login.UserName) || string.IsNullOrEmpty(login.Password))
            {
                return BadRequest("Invalid login data");
            }

            // Find the user by username and password
            var user = await _context.Users
                .FirstOrDefaultAsync(u => u.UserName == login.UserName && u.Password == login.Password);

            if (user == null)
            {
                return Unauthorized("Invalid username or password");
            }

            // Optional: Return user details (excluding sensitive data like password)
            var userDto = new UserLoginDto
            {
                //Id = user.Id,
                UserName = user.UserName,
                Password = user.Password,
                //Fname = user.Fname,
                //Lname = user.Lname,
                //Email = user.Email
            };

            return Ok(new { message = "Login successful", user = userDto });
        }
    }
}

namespace Expense_Tracker.Models
{
    public class User
    {
        public int Id { get; set; }
        public string? Fname { get; set; }
        public string? Lname { get; set; }
        public string? UserName { get; set; }
        public string? Password { get; set; }
        public string? Email { get; set; }
        public string? Phoneno { get; set; }
        public string? RoleID { get; set; }
        public string? UID { get; set; }
        public int? Income { get; set; }
        public string? Designation { get; set; }
        public string? Address { get; set; }
        public string? City { get; set; }
        public string? State { get; set; }
        public bool IsActive { get; set; }
        public bool IsDisabled { get; set; }
    }
}

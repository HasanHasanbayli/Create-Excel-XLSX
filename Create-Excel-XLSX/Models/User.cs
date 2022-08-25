namespace Create_Excel_XLSX.Models;

public class User
{
    public int Id { get; set; }
    public string? FirstName { get; set; }
    public string? LastName { get; set; }

    public IEnumerable<User> GetAllUsers()
    {
        return new List<User>
        {
            new()
            {
                Id = 1,
                FirstName = "Hasan",
                LastName = "Hasanbayli"
            },
            new()
            {
                Id = 2,
                FirstName = "Hasan2",
                LastName = "Hasanbayli2"
            },
            new()
            {
                Id = 3,
                FirstName = "Hasan3",
                LastName = "Hasanbayli3"
            },
            new()
            {
                Id = 4,
                FirstName = "Hasan4",
                LastName = "Hasanbayli4"
            }
        };
    }
}
using Create_Excel_XLSX;
using Create_Excel_XLSX.Models;

User user = new();
var users = user.GetAllUsers();

ExportOperation operation = new();
await operation.GenerateExcel(users);
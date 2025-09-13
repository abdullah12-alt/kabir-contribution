using System.Data;
using Microsoft.Data.SqlClient;
using System.Security.Cryptography;
using System.Text;
using static Server.Shared.Constants;
namespace Server.Services;


public interface IAuthenticationService
{
    public bool AuthenticateUser(string userId, string password);
    public void ChangePassword(string userId, string oldPassword, string newPassword, string confirmPassword);

}
public class AuthenticationService : IAuthenticationService
{


    private readonly string? _connectionString;

    public AuthenticationService(IConfiguration configuration)
    {
        _connectionString = configuration.GetConnectionString("DDSConnection");
    }

    public bool AuthenticateUser(string userId, string password)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            connection.Open();
            var query = "SELECT * FROM DD_USER WHERE USER_ID = @UserId";
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@UserId", userId);
                using (var reader = command.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        var recordStatus = reader["RECORD_STATUS"].ToString();
                        if (recordStatus == "I")
                        {
                            throw new Exception(Messages.UserInactive);
                        }

                        var storedPassword = reader["PSWTEXT"].ToString();
                        var hashedPassword = ComputeSHA256Hash(password);

                        if (storedPassword == hashedPassword)
                        {
                            UpdateLoginStats(userId, "L");
                            return true;
                        }
                        else
                        {
                            UpdateLoginStats(userId, "F");
                            throw new Exception(Messages.InvalidCredentials);
                        }
                    }
                    else
                    {
                        throw new Exception(Messages.UserNotExist);
                    }
                }
            }
        }
    }

    public void ChangePassword(string userId, string oldPassword, string newPassword, string confirmPassword)
    {
        if (newPassword.Length < 6)
        {
            throw new Exception("New password must be at least 6 characters long.");
        }

        if (newPassword != confirmPassword)
        {
            throw new Exception("New password and confirm password do not match.");
        }

        if (oldPassword == newPassword)
        {
            throw new Exception("New password cannot be the same as the old password.");
        }

        using (var connection = new SqlConnection(_connectionString))
        {
            connection.Open();
            var query = "SELECT PSWTEXT FROM DD_USER WHERE USER_ID = @UserId";
            using (var command = new SqlCommand(query, connection))
            {
                command.Parameters.AddWithValue("@UserId", userId);
                var storedPassword = command.ExecuteScalar()?.ToString();

                var hashedOldPassword = ComputeSHA256Hash(userId + oldPassword + "P#Ssa(fC");
                if (storedPassword != hashedOldPassword)
                {
                    throw new Exception("Old password is incorrect.");
                }

                var hashedNewPassword = ComputeSHA256Hash(userId + newPassword + "P#Ssa(fC");
                UpdatePassword(userId, hashedNewPassword);
            }
        }
    }

    private void UpdateLoginStats(string userId, string updateMode)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            connection.Open();
            var command = new SqlCommand("up_Login_Stats", connection)
            {
                CommandType = CommandType.StoredProcedure
            };
            command.Parameters.AddWithValue("@user_id", userId);
            command.Parameters.AddWithValue("@update_mode", updateMode);
            command.Parameters.AddWithValue("@misc_text", "");
            command.Parameters.AddWithValue("@version", "2.4");

            command.ExecuteNonQuery();
        }
    }

    private void UpdatePassword(string userId, string newPassword)
    {
        using (var connection = new SqlConnection(_connectionString))
        {
            connection.Open();
            var command = new SqlCommand("up_Login_Stats", connection)
            {
                CommandType = CommandType.StoredProcedure
            };
            command.Parameters.AddWithValue("@user_id", userId);
            command.Parameters.AddWithValue("@update_mode", "P");
            command.Parameters.AddWithValue("@misc_text", newPassword);
            command.Parameters.AddWithValue("@version", "2.4");

            command.ExecuteNonQuery();
        }
    }

    private string ComputeSHA256Hash(string input)
    {
        using (var sha256 = SHA256.Create())
        {
            var bytes = sha256.ComputeHash(Encoding.UTF8.GetBytes(input));
            var builder = new StringBuilder();
            foreach (var b in bytes)
            {
                builder.Append(b.ToString("x2"));
            }
            return builder.ToString();
        }
    }
}
# DDS Users Backend Implementation

This is the backend implementation for the DDS Users functionality, converted from the original VB.NET form (`frmDDSUsers.frm`) to a modern C# ASP.NET Core Web API.

## Overview

The DDS Users system allows administrators to manage user accounts in the Direct Deposit Subsystem (DDS). This includes creating, updating, deleting, and viewing user information with proper security and validation.

## Architecture

The implementation follows a clean architecture pattern with:

- **Models**: Data Transfer Objects (DTOs) for API requests and responses
- **Services**: Business logic layer with interface-based design
- **Controllers**: API endpoints for HTTP operations
- **Dependency Injection**: Service registration and lifecycle management

## Key Features

### User Management
- ✅ Create new users with password validation
- ✅ Update existing user information
- ✅ Delete users (with cascade deletion of related records)
- ✅ View all users or specific user by ID
- ✅ Check if user exists
- ✅ Validate user passwords

### Security Features
- ✅ Password hashing using SHA256 with salt
- ✅ Password confirmation validation
- ✅ Minimum password length enforcement (6 characters)
- ✅ User status management (Active/Inactive)

### Data Validation
- ✅ Required field validation
- ✅ Password strength requirements
- ✅ User ID uniqueness validation
- ✅ Input sanitization and trimming

## API Endpoints

### GET /api/ddsusers
Retrieves all DDS users ordered by user ID.

### GET /api/ddsusers/{userId}
Retrieves a specific DDS user by ID.

### POST /api/ddsusers
Creates a new DDS user.

**Request Body:**
```json
{
  "userId": "string",
  "userLastName": "string",
  "userFirstName": "string",
  "password": "string",
  "passwordConfirm": "string",
  "isDisabled": false
}
```

### PUT /api/ddsusers/{userId}
Updates an existing DDS user.

**Request Body:**
```json
{
  "userLastName": "string",
  "userFirstName": "string",
  "password": "string (optional)",
  "passwordConfirm": "string (optional)",
  "isDisabled": false
}
```

### DELETE /api/ddsusers/{userId}
Deletes a DDS user and all related records.

### GET /api/ddsusers/{userId}/exists
Checks if a user exists by ID.

### POST /api/ddsusers/{userId}/validate-password
Validates a user's password.

## Database Schema

The implementation works with the following database tables:

### DD_USER
- `USER_ID` (Primary Key) - User login ID
- `USER_LAST_NAME` - User's last name
- `USER_FIRST_NAME` - User's first name
- `PSWTEXT` - Hashed password
- `RECORD_STATUS` - User status (A=Active, I=Inactive)
- `CREATED_BY` - User who created the record
- `CREATED_DATETIME` - Creation timestamp
- `LAST_MOD_BY` - User who last modified the record
- `LAST_MOD_DATETIME` - Last modification timestamp

### Related Tables
- `DD_USER_ROLES` - User role assignments
- `DD_INSTITUTION_USER` - User institution access rights

## Password Security

Passwords are hashed using SHA256 with the following pattern:
```
Hash = SHA256(userId + password + "P#Ssa(fC")
```

This matches the original VB.NET implementation's password hashing mechanism.

## Error Handling

The API includes comprehensive error handling:

- **400 Bad Request**: Invalid input data or validation errors
- **404 Not Found**: User not found
- **500 Internal Server Error**: Unexpected server errors

All errors are logged using the application's logging infrastructure.

## Dependencies

- **Dapper**: Micro ORM for database operations
- **Microsoft.Data.SqlClient**: SQL Server database connectivity
- **ASP.NET Core**: Web framework
- **System.Security.Cryptography**: Password hashing

## Configuration

The service requires the following configuration:

1. **Database Connection**: Configured through `DapperDbContext`
2. **Logging**: Integrated with ASP.NET Core logging
3. **Dependency Injection**: Services registered in `Program.cs`

## Usage Examples

### Creating a New User
```csharp
var request = new CreateDdsUserRequest
{
    UserId = "JDOE",
    UserLastName = "Doe",
    UserFirstName = "John",
    Password = "securepassword123",
    PasswordConfirm = "securepassword123",
    IsDisabled = false
};

var user = await _ddsUserService.CreateUserAsync(request, "SYSTEM");
```

### Updating a User
```csharp
var request = new UpdateDdsUserRequest
{
    UserLastName = "Smith",
    UserFirstName = "Jane",
    IsDisabled = false
};

var updatedUser = await _ddsUserService.UpdateUserAsync("JDOE", request, "SYSTEM");
```

## Migration Notes

This implementation maintains compatibility with the original VB.NET system:

1. **Password Hashing**: Uses the same hashing algorithm and salt
2. **Database Schema**: Works with existing DD_USER table structure
3. **Business Rules**: Implements the same validation and business logic
4. **Data Integrity**: Maintains referential integrity with related tables

## Future Enhancements

Potential improvements for future versions:

1. **Authentication Integration**: Integrate with existing authentication system
2. **Role Management**: Add user role assignment functionality
3. **Institution Access**: Add institution access management
4. **Audit Logging**: Enhanced audit trail for user changes
5. **Password Policies**: Configurable password strength requirements
6. **Bulk Operations**: Support for bulk user operations 
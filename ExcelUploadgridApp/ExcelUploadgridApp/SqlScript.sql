  
-- Drop tables if they exist
IF OBJECT_ID('valuedetails', 'U') IS NOT NULL
    DROP TABLE valuedetails;
IF OBJECT_ID('columndetail', 'U') IS NOT NULL
    DROP TABLE columndetail;
IF OBJECT_ID('parentdetail', 'U') IS NOT NULL
    DROP TABLE parentdetail;
IF OBJECT_ID('ValueDetailsAudit', 'U') IS NOT NULL
    DROP TABLE ValueDetailsAudit;

-- Create parentdetail table to store datasource details
CREATE TABLE parentdetail (
    id INT PRIMARY KEY IDENTITY(1,1),
    datasource_name NVARCHAR(255) NOT NULL,
    description NVARCHAR(500),
    is_active BIT NOT NULL DEFAULT 1
);

-- Create columndetail table to store column details of each datasource
CREATE TABLE columndetail (
    id INT PRIMARY KEY IDENTITY(1,1),
    parent_id INT NOT NULL FOREIGN KEY REFERENCES parentdetail(id),
    column_name NVARCHAR(255) NOT NULL,
    data_type NVARCHAR(50) NOT NULL,
    is_required BIT NOT NULL,
    is_nullable BIT NOT NULL,
    screen_sequence INT,
    user_friendly_name NVARCHAR(255),
    display_format NVARCHAR(100),
    is_editable BIT NOT NULL DEFAULT 1,
    constraint_expression NVARCHAR(255),
    start_date DATETIME,
    end_date DATETIME,
    error_message NVARCHAR(500)
);

-- Create valuedetails table to store values for each column of each datasource
CREATE TABLE valuedetails (
    id INT PRIMARY KEY IDENTITY(1,1),
    column_id INT NOT NULL FOREIGN KEY REFERENCES columndetail(id),
    row_id INT NOT NULL,
    value NVARCHAR(MAX) NULL
);

-- Create ValueDetailsAudit table to store audit information
CREATE TABLE ValueDetailsAudit (
    AuditId INT PRIMARY KEY IDENTITY(1,1),
    ColumnId INT NOT NULL FOREIGN KEY REFERENCES columndetail(id),
    RowId INT NOT NULL,
    OldValue NVARCHAR(MAX),
    NewValue NVARCHAR(MAX),
    ColumnName NVARCHAR(MAX),
    DatasourceName NVARCHAR(MAX),
    ModifiedDate DATETIME DEFAULT GETDATE(),
    ModifiedBy NVARCHAR(255)
);

-- Insert data into parentdetail table
INSERT INTO parentdetail (datasource_name, description, is_active) VALUES 
('Datasource 1', 'Description for Datasource 1', 1),
('Datasource 2', 'Description for Datasource 2', 1);

-- Insert data into columndetail table
-- For Datasource 1
INSERT INTO columndetail (parent_id, column_name, data_type, is_required, is_nullable, screen_sequence, user_friendly_name, display_format, is_editable, constraint_expression, start_date, end_date, error_message) VALUES 
(1, 'client_id', 'INT', 1, 0, 1, 'Client ID', NULL, 1, 'value > 100', NULL, NULL, 'Client ID must be greater than 100'),
(1, 'address', 'NVARCHAR(255)', 1, 0, 2, 'Address', NULL, 1, NULL, NULL, NULL, 'Address must not be empty'),
(1, 'on_boarding_date', 'DATETIME', 0, 1, 3, 'On Boarding Date', NULL, 1, 'value <= DateTime.Now', NULL, NULL, 'On Boarding Date must be in the past'),
(1, 'checked_?', 'BIT', 1, 0, 4, 'Checked ?', NULL, 1, 'value == 0 || value == 1', NULL, NULL, 'Checked value must be 0 or 1');

-- For Datasource 2
-- Insert data into columndetail table for Datasource 2
INSERT INTO columndetail (parent_id, column_name, data_type, is_required, is_nullable, screen_sequence, user_friendly_name, display_format, is_editable, constraint_expression, start_date, end_date, error_message) VALUES 
(2, 'Name', 'NVARCHAR(255)', 1, 0, 1, 'Name', NULL, 1, NULL, NULL, NULL, 'Name must not be empty'),
(2, 'Age', 'INT', 0, 1, 2, 'Age', NULL, 1, 'value > 30', NULL, NULL, 'Age must be greater than 0'),
(2, 'Joining Date', 'DATETIME', 1, 0, 3, 'Joining Date', NULL, 1, 'value <= DateTime.Now', NULL, NULL, 'Joining Date must be in the past'),
(2, 'Salary', 'DECIMAL(10, 2)', 1, 0, 4, 'Salary', NULL, 1, NULL, NULL, NULL, 'Salary must not be empty'),
(2, 'Position', 'NVARCHAR(100)', 1, 0, 5, 'Position', NULL, 1, NULL, NULL, NULL, 'Position must not be empty'),
(2, 'Department', 'NVARCHAR(100)', 1, 0, 6, 'Department', NULL, 1, NULL, NULL, NULL, 'Department must not be empty'),
(2, 'Email', 'NVARCHAR(255)', 1, 1, 7, 'Email', NULL, 1, NULL, NULL, NULL, 'Email can be empty'),
(2, 'Phone', 'NVARCHAR(20)', 1, 1, 8, 'Phone', NULL, 1, NULL, NULL, NULL, 'Phone can be empty'),
(2, 'Address', 'NVARCHAR(255)', 1, 1, 9, 'Address', NULL, 1, NULL, NULL, NULL, 'Address can be empty'),
(2, 'City', 'NVARCHAR(100)', 1, 1, 10, 'City', NULL, 1, NULL, NULL, NULL, 'City can be empty');


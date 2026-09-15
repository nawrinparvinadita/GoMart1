/* =========================================================
   Go Mart Database - Microsoft SQL Server 2022 Setup Script
   =========================================================
   এই স্ক্রিপ্টটি SSMS (SQL Server Management Studio) অথবা
   Azure Data Studio-তে ওপেন করে পুরোটা Execute (F5) করলেই
   GoMartDB ডাটাবেস + সব টেবিল + সব Stored Procedure + Demo
   লগইন ডেটা তৈরি হয়ে যাবে।

   এই স্ক্রিপ্ট বারবার রান করলেও সমস্যা নেই (idempotent) -
   আগে থেকে থাকলে আবার তৈরি করবে না।

   GoMartApplication.csproj এর প্রতিটা ফর্মের SqlCommand যাচাই
   করে এই স্ক্রিপ্ট বানানো হয়েছে, তাই টেবিল/প্রসিডিউর/প্যারামিটারের
   নাম হুবহু কোডের সাথে মিলবে।
   ========================================================= */

IF NOT EXISTS (SELECT name FROM sys.databases WHERE name = N'GoMartDB')
BEGIN
    CREATE DATABASE GoMartDB;
END
GO

USE GoMartDB;
GO

-- =========================================================
-- TABLES
-- =========================================================

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'tblAdmin')
BEGIN
    CREATE TABLE tblAdmin
    (
        AdminID    NVARCHAR(50) PRIMARY KEY,
        [Password] NVARCHAR(50),
        FullName   NVARCHAR(50)
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'tblSeller')
BEGIN
    CREATE TABLE tblSeller
    (
        SellerID    INT IDENTITY(1,1) PRIMARY KEY,
        SellerName  NVARCHAR(50) UNIQUE,
        SellerAge   INT,
        SellerPhone NVARCHAR(50),   -- original script had NVARCHAR(10); widened so full BD phone numbers (e.g. 01XXXXXXXXX) don't get truncated
        SellerPass  NVARCHAR(50)
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'tblCategory')
BEGIN
    CREATE TABLE tblCategory
    (
        CatID        INT IDENTITY(1,1) PRIMARY KEY NOT NULL,
        CategoryName NVARCHAR(50),
        CategoryDesc NVARCHAR(50)
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'tblProduct')
BEGIN
    CREATE TABLE tblProduct
    (
        ProdID    INT IDENTITY(1,1) PRIMARY KEY NOT NULL,
        ProdName  NVARCHAR(50),
        ProdCatID INT,
        ProdPrice DECIMAL(10,2),
        ProdQty   INT
    );
END
GO

IF NOT EXISTS (SELECT * FROM sys.tables WHERE name = 'tblBill')
BEGIN
    CREATE TABLE tblBill
    (
        Bill_ID  INT PRIMARY KEY,
        SellerID NVARCHAR(50),
        SellDate NVARCHAR(50),
        TotalAmt DECIMAL(18,2)
    );
END
GO

-- =========================================================
-- STORED PROCEDURES: Admin  (used by AddAdmin.cs, Form1.cs)
-- =========================================================

CREATE OR ALTER PROCEDURE spAddAdmin
    @AdminID   NVARCHAR(50),
    @Password  NVARCHAR(50),
    @FullName  NVARCHAR(50)
AS
BEGIN
    INSERT INTO tblAdmin(AdminID, [Password], FullName)
    VALUES (@AdminID, @Password, @FullName);
END
GO

CREATE OR ALTER PROCEDURE spUpdateAdmin
    @AdminID   NVARCHAR(50),
    @Password  NVARCHAR(50),
    @FullName  NVARCHAR(50)
AS
BEGIN
    UPDATE tblAdmin
    SET [Password] = @Password, FullName = @FullName
    WHERE AdminID = @AdminID;
END
GO

CREATE OR ALTER PROCEDURE spDeleteAdmin
    @AdminID NVARCHAR(50)
AS
BEGIN
    DELETE FROM tblAdmin WHERE AdminID = @AdminID;
END
GO

-- Not called directly by the compiled app (Form1.cs uses an inline SELECT for login),
-- kept here for completeness / manual testing in SSMS.
CREATE OR ALTER PROCEDURE spAdminLogin
    @AdminID  NVARCHAR(50),
    @Password NVARCHAR(50)
AS
BEGIN
    SELECT TOP 1 AdminID, [Password], FullName
    FROM tblAdmin
    WHERE AdminID = @AdminID AND [Password] = @Password;
END
GO

-- =========================================================
-- STORED PROCEDURES: Seller  (used by frmAddNewSeller.cs)
-- =========================================================

CREATE OR ALTER PROCEDURE spSellerInsert
    @SellerName  NVARCHAR(50),
    @SellerAge   INT,
    @SellerPhone NVARCHAR(50),
    @SellerPass  NVARCHAR(50)
AS
BEGIN
    INSERT INTO tblSeller(SellerName, SellerAge, SellerPhone, SellerPass)
    VALUES (@SellerName, @SellerAge, @SellerPhone, @SellerPass);
END
GO

-- NOTE: name kept as "spSellerUpadte" (typo) on purpose -
-- the compiled app calls it by this exact name.
CREATE OR ALTER PROCEDURE spSellerUpadte
    @SellerID    INT,
    @SellerName  NVARCHAR(50),
    @SellerAge   INT,
    @SellerPhone NVARCHAR(50),
    @SellerPass  NVARCHAR(50)
AS
BEGIN
    UPDATE tblSeller
    SET SellerName = @SellerName, SellerAge = @SellerAge,
        SellerPhone = @SellerPhone, SellerPass = @SellerPass
    WHERE SellerID = @SellerID;
END
GO

CREATE OR ALTER PROCEDURE spSellerDelete
    @SellerID INT
AS
BEGIN
    DELETE FROM tblSeller WHERE SellerID = @SellerID;
END
GO

-- Not called directly by the compiled app (Form1.cs uses an inline SELECT for login),
-- kept here for completeness / manual testing in SSMS.
CREATE OR ALTER PROCEDURE spSellerLogin
    @SellerName NVARCHAR(50),
    @SellerPass NVARCHAR(50)
AS
BEGIN
    SELECT TOP 1 SellerName, SellerPass
    FROM tblSeller
    WHERE SellerName = @SellerName AND SellerPass = @SellerPass;
END
GO

-- =========================================================
-- STORED PROCEDURES: Category  (used by frmCategory.cs, AddProduct.cs, SellingForm.cs)
-- =========================================================

CREATE OR ALTER PROCEDURE spCatInsert
    @CategoryName NVARCHAR(50),
    @CategoryDesc NVARCHAR(50)
AS
BEGIN
    INSERT INTO tblCategory(CategoryName, CategoryDesc)
    VALUES (@CategoryName, @CategoryDesc);
END
GO

CREATE OR ALTER PROCEDURE spCatUpdate
    @CatID        INT,
    @CategoryName NVARCHAR(50),
    @CategoryDesc NVARCHAR(50)
AS
BEGIN
    UPDATE tblCategory
    SET CategoryName = @CategoryName, CategoryDesc = @CategoryDesc
    WHERE CatID = @CatID;
END
GO

CREATE OR ALTER PROCEDURE spCatDelete
    @CatID INT
AS
BEGIN
    DELETE FROM tblCategory WHERE CatID = @CatID;
END
GO

-- Not called directly by the compiled app (frmCategory.cs uses an inline SELECT to bind the grid),
-- kept here for completeness / manual testing in SSMS.
CREATE OR ALTER PROCEDURE spGetCatList
AS
BEGIN
    SELECT CatID AS CategoryID, CategoryName, CategoryDesc AS CategoryDescription
    FROM tblCategory;
END
GO

-- Used everywhere a category dropdown is populated (AddProduct.cs, SellingForm.cs)
CREATE OR ALTER PROCEDURE spGetCategory
AS
BEGIN
    SET NOCOUNT ON;
    SELECT CatID, CategoryName
    FROM tblCategory
    ORDER BY CategoryName ASC;
END
GO

-- =========================================================
-- STORED PROCEDURES: Product  (used by AddProduct.cs, SellingForm.cs)
-- =========================================================

CREATE OR ALTER PROCEDURE spCheckDuplicateProduct
    @ProdName  NVARCHAR(50),
    @ProdCatID INT
AS
BEGIN
    SET NOCOUNT ON;
    SELECT ProdName FROM tblProduct
    WHERE ProdName = @ProdName AND ProdCatID = @ProdCatID;
END
GO

CREATE OR ALTER PROCEDURE spInsertProduct
    @ProdName  NVARCHAR(50),
    @ProdCatID INT,
    @ProdPrice DECIMAL(10,2),
    @ProdQty   INT
AS
BEGIN
    INSERT INTO tblProduct(ProdName, ProdCatID, ProdPrice, ProdQty)
    VALUES (@ProdName, @ProdCatID, @ProdPrice, @ProdQty);
END
GO

CREATE OR ALTER PROCEDURE spUpdateProduct
    @ProdID    INT,
    @ProdName  NVARCHAR(50),
    @ProdCatID INT,
    @ProdPrice DECIMAL(10,2),
    @ProdQty   INT
AS
BEGIN
    UPDATE tblProduct
    SET ProdName = @ProdName, ProdCatID = @ProdCatID,
        ProdPrice = @ProdPrice, ProdQty = @ProdQty
    WHERE ProdID = @ProdID;
END
GO

CREATE OR ALTER PROCEDURE spDeleteProduct
    @ProdID INT
AS
BEGIN
    DELETE FROM tblProduct WHERE ProdID = @ProdID;
END
GO

CREATE OR ALTER PROCEDURE spGetAllProductList
AS
BEGIN
    SET NOCOUNT ON;
    SELECT t1.ProdID, t1.ProdName, t2.CategoryName,
           t1.ProdCatID AS CategoryID, t1.ProdPrice, t1.ProdQty
    FROM tblProduct AS t1
    INNER JOIN tblCategory AS t2 ON t1.ProdCatID = t2.CatID
    ORDER BY t1.ProdName, t2.CategoryName ASC;
END
GO

CREATE OR ALTER PROCEDURE spGetAllProductList_SearchByCat
    @ProdCatID INT
AS
BEGIN
    SET NOCOUNT ON;
    SELECT t1.ProdID, t1.ProdName, t2.CategoryName,
           t1.ProdCatID AS CategoryID, t1.ProdPrice, t1.ProdQty
    FROM tblProduct AS t1
    INNER JOIN tblCategory AS t2 ON t1.ProdCatID = t2.CatID
    WHERE t1.ProdCatID = @ProdCatID
    ORDER BY t1.ProdName, t2.CategoryName ASC;
END
GO

-- =========================================================
-- STORED PROCEDURES: Bill  (used by SellingForm.cs)
-- =========================================================

CREATE OR ALTER PROCEDURE spInsertBill
    @Bill_ID  INT,
    @SellerID NVARCHAR(50),
    @SellDate NVARCHAR(50),
    @TotalAmt DECIMAL(18,2)
AS
BEGIN
    INSERT INTO tblBill(Bill_ID, SellerID, SellDate, TotalAmt)
    VALUES (@Bill_ID, @SellerID, @SellDate, @TotalAmt);
END
GO

CREATE OR ALTER PROCEDURE spGetBillList
AS
BEGIN
    SET NOCOUNT ON;
    SELECT Bill_ID, SellerID, SellDate, TotalAmt
    FROM tblBill
    ORDER BY Bill_ID DESC;
END
GO

-- =========================================================
-- SEED / DEMO DATA
-- 'Coder' / '12345' - Form1.cs (Form1_Load) নিজে থেকেই এই
-- ইউজারনেম-পাসওয়ার্ড লগইন বক্সে ভরে রাখে, তাই এটা এই স্ক্রিপ্টেও রাখা হলো।
-- =========================================================

IF NOT EXISTS (SELECT 1 FROM tblAdmin WHERE AdminID = 'Coder')
BEGIN
    INSERT INTO tblAdmin (AdminID, [Password], FullName)
    VALUES ('Coder', '12345', 'Demo Administrator');
END
GO

IF NOT EXISTS (SELECT 1 FROM tblCategory WHERE CategoryName = 'Groceries')
BEGIN
    INSERT INTO tblCategory (CategoryName, CategoryDesc)
    VALUES ('Groceries', 'Everyday grocery items');
END
GO

IF NOT EXISTS (SELECT 1 FROM tblCategory WHERE CategoryName = 'Beverages')
BEGIN
    INSERT INTO tblCategory (CategoryName, CategoryDesc)
    VALUES ('Beverages', 'Drinks and refreshments');
END
GO

-- =========================================================
-- Verify (optional - just to see the result after running)
-- =========================================================
SELECT * FROM tblAdmin;
SELECT * FROM tblCategory;
SELECT * FROM tblSeller;
SELECT * FROM tblProduct;
SELECT * FROM tblBill;
GO

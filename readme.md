Overseas Import Corporation - E-Commerce System
Overview
ASP Classic-based e-commerce platform for automotive parts distribution with multi-language support, tiered pricing, and advanced order management.
---
🚀 Key Features
1. User Management
•	Role-based access control with 4 user levels (1, 3, 4, 6)
•	Multi-user support for business accounts (B001 + client_multi)
•	Session-based authentication
•	Password change enforcement (default: 11111)
•	Tax configuration per user (city/state tax)
•	Custom discount rates per account
2. Product Catalog
•	Part search by:
•	Part number
•	Family/Category
•	Model/Keywords
•	Special items filter
•	Replacement part suggestions
•	Availability display (50+ shown as "50+")
•	Multi-language descriptions (English/Spanish)
•	Image support (JPG/PNG/GIF)
•	Price tiers based on user level:
•	Level 1: field_3
•	Level 3: field_4
•	Level 4: field_7
•	Level 6: field_8
3. Shopping Cart
•	Ajax-powered add/update/delete
•	Real-time quantity updates
•	Price display:
•	Regular prices
•	Special prices (red)
•	Liquidation prices (red "L")
•	Discount calculations
•	Tax calculation:
•	City Tax: 1%
•	State Tax: 10.5%
•	Multi-part orders (splits at 22 items per order number)
4. Order Processing
•	Order number generation via stored procedure GetNextOrderNumber
•	Concurrency control with retry logic
•	Delivery types:
•	S - Same Day (until 9:50 AM)
•	D - Next Day
•	P - Pickup
•	Instructions field (25 char max, split into multiple comments)
•	Inventory deduction on order submit
•	Order splitting for:
•	Regular items vs. special items
•	Orders > 22 parts
5. Account Management
•	Order history with:
•	Order number
•	Date/time
•	Total amount
•	Detailed line items
•	Account statements:
•	Invoice amounts/dates
•	Payment receipts
•	Discounts
•	Running balance
•	Aging buckets (Current, 30, 60, 90, 120+ days)
•	Password-protected statements (stored in accountStatementPassword)
6. Multi-Language Support
•	English (E) / Spanish (S)
•	Session-based language selection
•	Translatable labels via Lang() function

📁 File Structure

/overseaspr/
├── index.asp                       # Login page
├── part_search.asp                 # Product search interface
├── cart.asp                        # Shopping cart page
├── cart_iframe.asp                 # Cart details (embedded)
├── cart_proceed.asp                # Order submission
├── order_detail.asp                # Order confirmation
├── account_statement_iframe.asp    # Account statement
├── change_password.asp             # Password update
├── procedures.asp                  # Core business logic
├── Connections/
│   └── overseaspr.asp             # DB connection string
├── images/                         # UI assets
│   ├── oiclogo2.gif
│   ├── rutas_same_day2.png
│   └── tarjetas2.png
└── parts_images/                   # Product photos
    └── {part_number}.jpg/png/gif

🗄️ Database Schema

Core Tables

users
- user_auto_id (PK)
- user_name
- user_pwd
- user_level (1,3,4,6,9)
- user_client_code
- discount (%)
- citytax (Y/N)
- statetax (Y/N)
- lang (E/S)
- samedayflag (Y/N)
- accountStatementPassword

clients_cart
- shop_auto_id (PK)
- shop_client_code
- shop_client_user
- shop_product_id
- shop_part_number
- shop_quantity
- shop_part_price
- shop_order_date
- client_multi
- shop_type (W=Web)

clients_orders
- order_id (PK)
- order_number (grouped)
- order_part
- order_client
- order_user
- order_qty
- order_date
- order_type (D/S/P)
- item_price
- comments
- deliv_type
- zcustomer
- order_status (O/P)

partmst1_distinct
- field_1 (part_number) PK
- field_2 (description_spanish)
- field_3 (price_level_1)
- field_4 (price_level_3)
- field_5 (quantity_on_hand)
- field_7 (price_level_4)
- field_8 (price_level_6)
- english_version
- replacement
- familia_descripcion
- family_description
- image_exist (0/1/2)

prespecials
- specials (part_number)
- especial (special_price)

prod_liqui
- field_1 (part_number)
- price (liquidation_price)
- regular (original_price)

next_order_number
- current_number
- in_use (Y/N)

invoice
- cust_inv (PK)
- cust (FK to users)
- inv (invoice_number)
- invamt
- inv_date
- payment_amt
- payment_date
- disc1, disc2

🔧 Stored Procedures
GetNextOrderNumber
CREATE PROCEDURE GetNextOrderNumber
    @NextOrderNum INT OUTPUT
AS
BEGIN
    SET NOCOUNT ON;
    BEGIN TRANSACTION;
    
    -- Lock row
    UPDATE next_order_number 
    SET in_use = 'Y'
    WHERE in_use = 'N';
    
    -- Get next number
    SELECT @NextOrderNum = current_number + 1
    FROM next_order_number;
    
    -- Update counter
    UPDATE next_order_number
    SET current_number = @NextOrderNum,
        in_use = 'N';
    
    COMMIT TRANSACTION;
END

🔐 Security Features
SQL Injection Protection
•	GetSecureVal() function sanitizes inputs
•	Replaces single quotes: ' → ''
•	Numeric validation

Session Management
Session("MM_Username")          ' User ID
Session("MM_UserAuthorization") ' Auth token
Session("user_level")           ' Access level
Session("MM_discount")          ' Discount %
Session("MM_CityTax")           ' Y/N
Session("MM_StateTax")          ' Y/N
Session("lang")                 ' E/S
Session("MM_UserId_Multi")      ' For B001 multi-accounts

Password Requirements
•	Default: 11111 (forces change on first login)
•	Stored in plain text (⚠️ Security Risk)

🛠️ Setup Instructions
Prerequisites
•	IIS 7.0+ with ASP Classic enabled
•	SQL Server 2008+
•	ODBC DSN configured
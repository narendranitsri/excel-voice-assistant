"""
generate_samples.py
===================
Generates 4 small sample Excel files in the /samples folder to test:
1. sample_orders.xlsx: VLOOKUP with Prices, multi-column sorting, and filtering
2. sample_prices.xlsx: Key matching (Item vs Item_Name), Unit_Price retrieval
3. sample_inventory.xlsx: Common items, differences, and warehouse stock
4. sample_messy_data.xlsx: Test Data Quality Report, auto-fix, duplicate rows, blanks, and fuzzy city standardisation
"""

import os
import pandas as pd

def generate_samples():
    samples_dir = "samples"
    os.makedirs(samples_dir, exist_ok=True)
    print(f"Generating sample Excel files in '{samples_dir}/'...")

    # 1. Orders
    df_orders = pd.DataFrame({
        "Order_ID": [101, 102, 103, 104, 105, 106, 107, 108, 109, 110],
        "Item": [
            "Laptop Dell XPS", "Wireless Mouse", "Mechanical Keyboard", "USB-C Hub", 
            "Laptop Dell XPS", "Monitor 27-inch", "Desk Lamp LED", "Wireless Mouse", 
            "Ergonomic Chair", "Noise Cancelling Headphones"
        ],
        "Customer": [
            "Rajesh Sharma", "Priya Patel", "Amit Kumar", "Sneha Gupta", 
            "Vikram Malhotra", "Rajesh Sharma", "Ananya Sen", "Karan Johar", 
            "Sunita Verma", "Deepak Rao"
        ],
        "Quantity": [1, 3, 2, 5, 2, 1, 4, 2, 1, 2],
        "City": [
            "Mumbai", "Delhi", "Bengaluru", "Delhi", "Pune", 
            "Mumbai", "Kolkata", "Hyderabad", "Chennai", "Bengaluru"
        ],
        "Order_Date": [
            "2026-03-01", "2026-03-02", "2026-03-02", "2026-03-03", "2026-03-04",
            "2026-03-05", "2026-03-05", "2026-03-06", "2026-03-07", "2026-03-08"
        ]
    })
    orders_path = os.path.join(samples_dir, "sample_orders.xlsx")
    df_orders.to_excel(orders_path, index=False)
    print(f"✓ Created {orders_path}")

    # 2. Prices
    df_prices = pd.DataFrame({
        "Item_Name": [
            "Laptop Dell XPS", "Wireless Mouse", "Mechanical Keyboard", "USB-C Hub", 
            "Monitor 27-inch", "Desk Lamp LED", "Ergonomic Chair", "Noise Cancelling Headphones",
            "Webcam 4K Ultra", "Portable SSD 1TB"
        ],
        "Category": [
            "Computers", "Accessories", "Peripherals", "Accessories", 
            "Displays", "Office Decor", "Furniture", "Audio",
            "Cameras", "Storage"
        ],
        "Unit_Price": [74999, 899, 2999, 1499, 18999, 1299, 14500, 8999, 5499, 6999],
        "GST_Percent": [18, 18, 18, 18, 18, 12, 18, 18, 18, 18]
    })
    prices_path = os.path.join(samples_dir, "sample_prices.xlsx")
    df_prices.to_excel(prices_path, index=False)
    print(f"✓ Created {prices_path}")

    # 3. Inventory
    df_inventory = pd.DataFrame({
        "Product": [
            "Laptop Dell XPS", "Wireless Mouse", "Mechanical Keyboard", "Monitor 27-inch", 
            "Webcam 4K Ultra", "Graphics Tablet 10-inch", "Studio Microphone", "Cable Organizer Box"
        ],
        "Warehouse": [
            "Bhiwandi-West", "Gurugram-Central", "Bengaluru-Whitefield",
            "Bhiwandi-West", "Gurugram-Central", "Hyderabad-Gachibowli",
            "Bengaluru-Whitefield", "Bhiwandi-West"
        ],
        "Stock_Count": [15, 210, 85, 34, 18, 9, 22, 140],
        "Min_Reorder_Level": [10, 50, 20, 15, 10, 10, 15, 30]
    })
    inventory_path = os.path.join(samples_dir, "sample_inventory.xlsx")
    df_inventory.to_excel(inventory_path, index=False)
    print(f"✓ Created {inventory_path}")

    # 4. Messy Data (to test Data Quality Report & Auto-Fix)
    df_messy = pd.DataFrame({
        "Customer_Name": [
            "  amit sharma  ", "priya patel", "Rahul Verma", "Sneha Rao", 
            "priya patel", "Vikram Singh", None, "anita desai", "kunal gupta", "Rohit Mehra"
        ],
        "City": [
            "Dilli", "delhi ", "Delhi", "Bengaluru", 
            "delhi ", "Mumbai", "bombay", "Pune", "calcutta", "kolkata"
        ],
        "Phone": [
            "9820012345", "98200 12345", "+919820012345", "919820012345",
            "98200 12345", "98200", "9988776655", "invalid-phone", "9876543210", None
        ],
        "Email": [
            "amit@gmail.com", "priya@domain", "rahul.verma@corp.in", "sneha@gmail.com",
            "priya@domain", "bad@@email.com", "anita@yahoo.co.in", "contact@test", "kunal@outlook.com", "rohit@test.com"
        ],
        "Amount": [
            5000, 12000, 450, 18000, 
            12000, 999999, 3400, 8900, -500, 15000
        ],
        "Registration_Date": [
            "01-03-2026", "2026/03/02", "03-03-2026", "04/03/2026",
            "2026/03/02", "2026-03-05", "06-03-2026", None, "08-03-2026", "09/03/2026"
        ]
    })
    messy_path = os.path.join(samples_dir, "sample_messy_data.xlsx")
    df_messy.to_excel(messy_path, index=False)
    print(f"✓ Created {messy_path} (Contains duplicates, blanks, spaces, city variants, and bad emails/phones)")

if __name__ == "__main__":
    generate_samples()

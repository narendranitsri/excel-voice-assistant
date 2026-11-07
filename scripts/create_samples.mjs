import * as XLSX from 'xlsx';
import * as fs from 'fs';
import * as path from 'path';

const samplesDir = "samples";
if (!fs.existsSync(samplesDir)) {
  fs.mkdirSync(samplesDir, { recursive: true });
}

// 1. Orders
const orders = [
  { Order_ID: 101, Item: "Laptop Dell XPS", Customer: "Rajesh Sharma", Quantity: 1, City: "Mumbai", Order_Date: "2026-03-01" },
  { Order_ID: 102, Item: "Wireless Mouse", Customer: "Priya Patel", Quantity: 3, City: "Delhi", Order_Date: "2026-03-02" },
  { Order_ID: 103, Item: "Mechanical Keyboard", Customer: "Amit Kumar", Quantity: 2, City: "Bengaluru", Order_Date: "2026-03-02" },
  { Order_ID: 104, Item: "USB-C Hub", Customer: "Sneha Gupta", Quantity: 5, City: "Delhi", Order_Date: "2026-03-03" },
  { Order_ID: 105, Item: "Laptop Dell XPS", Customer: "Vikram Malhotra", Quantity: 2, City: "Pune", Order_Date: "2026-03-04" },
  { Order_ID: 106, Item: "Monitor 27-inch", Customer: "Rajesh Sharma", Quantity: 1, City: "Mumbai", Order_Date: "2026-03-05" },
  { Order_ID: 107, Item: "Desk Lamp LED", Customer: "Ananya Sen", Quantity: 4, City: "Kolkata", Order_Date: "2026-03-05" },
  { Order_ID: 108, Item: "Wireless Mouse", Customer: "Karan Johar", Quantity: 2, City: "Hyderabad", Order_Date: "2026-03-06" },
  { Order_ID: 109, Item: "Ergonomic Chair", Customer: "Sunita Verma", Quantity: 1, City: "Chennai", Order_Date: "2026-03-07" },
  { Order_ID: 110, Item: "Noise Cancelling Headphones", Customer: "Deepak Rao", Quantity: 2, City: "Bengaluru", Order_Date: "2026-03-08" }
];
const wsOrders = XLSX.utils.json_to_sheet(orders);
const wbOrders = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wbOrders, wsOrders, "Orders");
XLSX.writeFile(wbOrders, path.join(samplesDir, "sample_orders.xlsx"));
XLSX.writeFile(wbOrders, "sample_orders.xlsx");

// 2. Prices
const prices = [
  { Item_Name: "Laptop Dell XPS", Category: "Computers", Unit_Price: 74999, GST_Percent: 18 },
  { Item_Name: "Wireless Mouse", Category: "Accessories", Unit_Price: 899, GST_Percent: 18 },
  { Item_Name: "Mechanical Keyboard", Category: "Peripherals", Unit_Price: 2999, GST_Percent: 18 },
  { Item_Name: "USB-C Hub", Category: "Accessories", Unit_Price: 1499, GST_Percent: 18 },
  { Item_Name: "Monitor 27-inch", Category: "Displays", Unit_Price: 18999, GST_Percent: 18 },
  { Item_Name: "Desk Lamp LED", Category: "Office Decor", Unit_Price: 1299, GST_Percent: 12 },
  { Item_Name: "Ergonomic Chair", Category: "Furniture", Unit_Price: 14500, GST_Percent: 18 },
  { Item_Name: "Noise Cancelling Headphones", Category: "Audio", Unit_Price: 8999, GST_Percent: 18 },
  { Item_Name: "Webcam 4K Ultra", Category: "Cameras", Unit_Price: 5499, GST_Percent: 18 },
  { Item_Name: "Portable SSD 1TB", Category: "Storage", Unit_Price: 6999, GST_Percent: 18 }
];
const wsPrices = XLSX.utils.json_to_sheet(prices);
const wbPrices = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wbPrices, wsPrices, "Prices");
XLSX.writeFile(wbPrices, path.join(samplesDir, "sample_prices.xlsx"));
XLSX.writeFile(wbPrices, "sample_prices.xlsx");

// 3. Inventory
const inventory = [
  { Product: "Laptop Dell XPS", Warehouse: "Bhiwandi-West", Stock_Count: 15, Min_Reorder_Level: 10 },
  { Product: "Wireless Mouse", Warehouse: "Gurugram-Central", Stock_Count: 210, Min_Reorder_Level: 50 },
  { Product: "Mechanical Keyboard", Warehouse: "Bengaluru-Whitefield", Stock_Count: 85, Min_Reorder_Level: 20 },
  { Product: "Monitor 27-inch", Warehouse: "Bhiwandi-West", Stock_Count: 34, Min_Reorder_Level: 15 },
  { Product: "Webcam 4K Ultra", Warehouse: "Gurugram-Central", Stock_Count: 18, Min_Reorder_Level: 10 },
  { Product: "Graphics Tablet 10-inch", Warehouse: "Hyderabad-Gachibowli", Stock_Count: 9, Min_Reorder_Level: 10 },
  { Product: "Studio Microphone", Warehouse: "Bengaluru-Whitefield", Stock_Count: 22, Min_Reorder_Level: 15 },
  { Product: "Cable Organizer Box", Warehouse: "Bhiwandi-West", Stock_Count: 140, Min_Reorder_Level: 30 }
];
const wsInventory = XLSX.utils.json_to_sheet(inventory);
const wbInventory = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wbInventory, wsInventory, "Inventory");
XLSX.writeFile(wbInventory, path.join(samplesDir, "sample_inventory.xlsx"));
XLSX.writeFile(wbInventory, "sample_inventory.xlsx");

// 4. Messy Data
const messy = [
  { Customer_Name: "  amit sharma  ", City: "Dilli", Phone: "9820012345", Email: "amit@gmail.com", Amount: 5000, Registration_Date: "01-03-2026" },
  { Customer_Name: "priya patel", City: "delhi ", Phone: "98200 12345", Email: "priya@domain", Amount: 12000, Registration_Date: "2026/03/02" },
  { Customer_Name: "Rahul Verma", City: "Delhi", Phone: "+919820012345", Email: "rahul.verma@corp.in", Amount: 450, Registration_Date: "03-03-2026" },
  { Customer_Name: "Sneha Rao", City: "Bengaluru", Phone: "919820012345", Email: "sneha@gmail.com", Amount: 18000, Registration_Date: "04/03/2026" },
  { Customer_Name: "priya patel", City: "delhi ", Phone: "98200 12345", Email: "priya@domain", Amount: 12000, Registration_Date: "2026/03/02" },
  { Customer_Name: "Vikram Singh", City: "Mumbai", Phone: "98200", Email: "bad@@email.com", Amount: 999999, Registration_Date: "2026-03-05" },
  { Customer_Name: null, City: "bombay", Phone: "9988776655", Email: "anita@yahoo.co.in", Amount: 3400, Registration_Date: "06-03-2026" },
  { Customer_Name: "anita desai", City: "Pune", Phone: "invalid-phone", Email: "contact@test", Amount: 8900, Registration_Date: null },
  { Customer_Name: "kunal gupta", City: "calcutta", Phone: "9876543210", Email: "kunal@outlook.com", Amount: -500, Registration_Date: "08-03-2026" },
  { Customer_Name: "Rohit Mehra", City: "kolkata", Phone: null, Email: "rohit@test.com", Amount: 15000, Registration_Date: "09/03/2026" }
];
const wsMessy = XLSX.utils.json_to_sheet(messy);
const wbMessy = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(wbMessy, wsMessy, "Messy_Data");
XLSX.writeFile(wbMessy, path.join(samplesDir, "sample_messy_data.xlsx"));

console.log("Created all 4 sample files in /samples folder!");

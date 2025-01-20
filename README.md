# MEDITECH Magic Purchase Order Creation Script

This project automates the creation of purchase orders in the MEDITECH Magic system, specifically targeting the MM module. It integrates with Excel files for input, uses Python for data handling, and interacts with the MEDITECH GUI to streamline order management.

## Features

- Automates the purchase order creation process within MEDITECH Magic MM.
- Reads purchase order details and line items from an Excel file.
- Simulates keypresses and interactions with the MEDITECH Magic GUI.
- Supports customization of order types, vendors, delivery dates, and more.

## Requirements

- Python 3.8 or later
- Libraries:
  - `ctypes` (built-in)
  - `openpyxl`
  - `tkinter` (built-in)

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/brandonhenness/MEDITECH.Create-Purchase-Order.git
   cd MEDITECH.Create-Purchase-Order
   ```
2. Install the required dependencies:
   ```bash
   pip install openpyxl
   ```

## Usage

1. Prepare an Excel file with purchase order details:
   - The workbook should contain two sheets:
     - `PurchaseOrderInfo`: High-level purchase order details.
     - `LineItems`: Details for individual line items.

2. Run the script:
   ```bash
   python create_purchase_order.py
   ```

3. Use the file dialog to select the Excel file.

4. The script will:
   - Parse the Excel file.
   - Interact with the MEDITECH Magic MM GUI to create the purchase order and line items.

## Excel File Format

### `PurchaseOrderInfo` Sheet
| Column                     | Description                           |
|----------------------------|---------------------------------------|
| Purchase Order Number      | The PO number or `N` for new.         |
| Purchase Order Type        | The type of order (e.g., PURCHASE).   |
| Delivery Date              | Delivery date (default: 'T' for today). |
| Vendor                    | Vendor information.                   |

### `LineItems` Sheet
| Column                   | Description                           |
|--------------------------|---------------------------------------|
| Item Number              | Unique item number.                  |
| Common Name              | Common name for the item.            |
| Quantity                 | Quantity of the item.                |

## License

MEDITECH Magic Purchase Order Creation Script is licensed under the [GNU General Public License v3.0](LICENSE).

## Contributing

Contributions are welcome! Feel free to submit issues or pull requests.

## Disclaimer

This script interacts with the MEDITECH Magic GUI. Ensure the application is running and accessible during the script's execution.


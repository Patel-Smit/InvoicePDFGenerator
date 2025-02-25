# Invoice Generator

#### Video Demo: <https://youtu.be/F8XG-WrIS30>

#### Description: Invoice Generator is a Python-based tool designed to streamline the process of generating invoices for small services-based business. This system helps stores manage customer services, generate invoices invoices in .docx and .pdf formats, and track sales data efficiently.

## Features

- **Service Management:** Add, list, and manage services offered by your store.
- **Shopping Cart:** Add, remove, and view services in a customer's cart.
- **Checkout:** Generate a detailed invoice for the customer with tax calculation.
- **Invoice Generation:** Automatically creates and saves invoices in PDF format.
- **Sales Data Tracking:** Logs sales data to a CSV file for easy tracking.
- **Customizable:** Business details and invoice numbers can be customized using a configuration file.

## Requirements

Before running the script, ensure you have the following dependencies installed:

- `python-docx`
- `tabulate`
- `docx2pdf`
- `configparser`
- `difflib`
- `csv`
- `datetime`
- `os`

You can install these dependencies using `pip`:

```bash
pip install python-docx tabulate docx2pdf configparser
```

Needs Microsoft Word and tested on Windows 11.

## Project Structure

- **`invoice_template.docx`**: A Word document template used for generating invoices.
- **`config.ini`**: A configuration file containing business information and invoice numbering.
- **`services_list.csv`**: A CSV file containing available services and their prices.
- **`sales_data.csv`**: A CSV file that stores sales data (created automatically).
- **`invoices/`**: A directory where generated PDF invoices are stored.
- **`temp_files/`**: A directory where generated temporary invoices word files are stored.
- **`project.py`**: The main script that runs the invoice generator.
- **`test_project.py`**: The script that test few methods from **'project.py'**.

## How It Works

1. **Customer Details**: The program prompts the user to enter customer details, including their name and address.
2. **Service Selection**: The user can select services to add to the cart, view the cart, and remove services if necessary.
3. **Checkout**: The user can proceed to checkout, where the invoice is generated.
4. **Invoice Generation**: The system creates an invoice using a Word document template, fills in the customer and transaction details, and converts the document to PDF format.
5. **Sales Data Logging**: All sales data, including customer details and total charges, is saved to `sales_data.csv` for record-keeping.

## Configuration

The `config.ini` file contains important configuration settings, including:

- **Business Information**: Customize your business name, address, contact information, and payment details.
- **Invoice Number**: The invoice number is auto-incremented after each transaction. The starting invoice number can be modified.

Example `config.ini` structure:

```ini
[settings]
invoice_number = 000001

[business_data]
business_name = Keshav Tech
business_street_address = 123 Tech Street
business_city_province = Tech City, TP
business_email_address = contact@keshavtech.test
business_contact = +1-123-456-7890
business_payment_recipient = Keshav Tech Ltd.
business_payment_bank = Tech Bank
business_payment_IBAN = XX1234567890
business_payment_BIC = TCHBANKX
```

## Methods Overview

### `main()`

Initializes the system, loads business data from a configuration file, greets the customer, and displays the main menu.

### `main_option_menu(customer_name, customer_address, cart)`

Displays the main menu and handles user input for various operations like shopping, viewing cart, and checkout.

### `load_available_services(services_list, services_name_list)`

Loads available services from a CSV file and stores them in `services_list` and `services_name_list`.

### `show_available_services(services_list)`

Prints the list of available services with their corresponding prices.

### `get_cx_details()`

Prompts the user for customer details like name, address, city, and postal code, with validation checks.

### `_help()`

Displays a help menu with options for the user to navigate the system.

### `shop(cx_name, cx_address, cart)`

Handles the shopping process where users can select services and add them to the cart.

### `add_to_cart(item, cart)`

Adds a selected service to the shopping cart. If the service already exists in the cart, it updates the quantity.

### `remove_from_cart(item_name, cart)`

Removes a specified service or a specified quantity from the shopping cart.

### `suggest_item(item, items)`

Suggests closely matching items if the user input does not exactly match any available services.

### `show_cart(cart)`

Displays the current contents of the shopping cart in a tabulated format along with the subtotal.

### `checkout(cx_name, cx_address, cart)`

Initiates the checkout process, confirms the purchase, and generates an invoice.

### `replace_text(paragraph, old_text, new_text)`

Replaces specified text in a paragraph of a Word document.

### `generate_invoice(cx_name, cx_address, cart)`

Generates an invoice document based on the customer details and cart items, and saves it as a PDF.

### `post_data(invoice_number, date_time, cx_name, cx_address, cart, total)`

Saves the sales data (including invoice number and cart details) to a CSV file for record-keeping.

## Usage

1. Run the `project.py` script:

   ```bash
   python project.py
   ```

2. The program will display the main menu with options to start shopping, view available services, check the cart, or proceed to checkout.

3. After the checkout process is complete, a PDF invoice will be generated and saved in the `invoices/` directory. The sales data will be logged in `sales_data.csv`.

## Error Handling

- The program ensures valid customer details and service selections through input validation.
- If a service is not found, the program suggests similar services using the `difflib` library.
- The script handles potential errors during file generation and input/output operations.

## Customization

- You can modify the invoice template by editing the `invoice_template.docx` file.
- Update business information and payment details in the `config.ini` file.
- Add or modify services in the `services_list.csv` file.

## Output Invoice Sample
![Invoice Sample.png](Invoice%20Sample.png)

## Acknowledgments

Special thanks to the libraries and tools used in this project: `docx`, `tabulate`, `docx2pdf`, and Python's standard libraries.

---

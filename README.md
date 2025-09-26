# Web Automation Suite

A comprehensive Python-based automation framework for web scraping and data collection using Selenium WebDriver.

## 🚀 Features

- **Unified Interface**: Single entry point for multiple automation tasks
- **VPN Integration**: Built-in VPN connection management for secure operations
- **Smart Download Management**: Automated file downloads with rename capabilities
- **Web Scraping**: Extract data from dashboards and web pages
- **Excel Export**: Automated summary generation in Excel format
- **Error Handling**: Robust error handling with task continuation
- **Modular Design**: Easy to extend with new tasks

## 📋 Prerequisites

- Python 3.8+
- Google Chrome browser
- ChromeDriver (compatible with your Chrome version)
- VPN client (optional, can be disabled)

## 🔧 Installation

1. Clone the repository:
git clone https://github.com/yourusername/web-automation-suite.git
cd web-automation-suite


2. Install required packages:
pip install -r requirements.txt


3. Download ChromeDriver from [here](https://chromedriver.chromium.org/) and ensure it's in your PATH.

## 📦 Dependencies

Create a `requirements.txt` file with:

selenium==4.15.0
pandas==2.1.3
openpyxl==3.1.2
python-dateutil==2.8.2


## ⚙️ Configuration

Update the following in the main script:
automation = UnifiedAutomation(
    login_url="your_login_url",
    login_id="your_username",
    password="your_password",
    use_vpn=True,  # Set to False if not using VPN
)


### VPN Configuration (Optional)
If using VPN, update these parameters:
vpn_exe=r"C:\Program Files\OpenVPN Connect\OpenVPNConnect.exe"
vpn_shortcut_id="your_vpn_shortcut_id"

## 🎯 Usage

### Basic Usage

Run the script:
python unified_automation.py


### Menu Options

When running the script, you'll see:

Enter the Login ID :
Enter the password : 

Available tasks:
1. Platform A - Download All Reports
2. Platform B - Download Report
3. Platform C - Scrape Metrics
4. Platform D - Scrape Metrics
5. Platform E - Scrape Row Count
6. Platform E - Download Transactions
7. Platform F - Scrape Metrics

Enter task numbers separated by commas (e.g., 1,3,5)
Or press Enter to run ALL tasks

Your selection:

### Examples

- Run specific tasks: `1,3,5`
- Run all tasks: Press `Enter`
- Run single task: `2`

## 📊 Output

The automation generates:

1. **Downloaded Files**: Saved to the script directory with standardized names
2. **Excel Summary**: `automation_summary.xlsx` containing:
   - Scraped data values
   - Downloaded file paths
   - Timestamps

### Output Structure
your-project/
├── unified_automation.py
├── automation_summary.xlsx
├── Platform_A_Report.xlsx
├── Platform_B_Report.xlsx
└── Platform_C_Transactions.xlsx

## 🏗️ Architecture

### Main Components

1. **UnifiedAutomation Class**
   - Central controller for all automation tasks
   - Manages browser lifecycle
   - Handles VPN connections

2. **Task Methods**
   - `_download_report()`: Handles file downloads
   - `_scrape_data()`: Extracts data from web pages
   - `_wait_for_download()`: Manages download completion

3. **Helper Methods**
   - `_login()`: Unified login handler
   - `_make_driver()`: WebDriver initialization
   - `save_results_to_excel()`: Export functionality

## 🔐 Security Notes

- Store credentials securely (consider environment variables)
- Use VPN for additional security
- Review and update login credentials regularly
- Don't commit sensitive data to version control

## 🛠️ Customization

### Adding New Tasks

1. Create a new method in the class:
def new_platform_scrape(self):
    """Scrape data from new platform"""
    # Your scraping logic here
    return results

2. Add to `tasks_config()`:
"8": {"name": "New Platform - Scrape Data", "method": self.new_platform_scrape, "type": "scrape"}

### Modifying Date Ranges

Update the date calculation methods:
- `_get_previous_month_dates()`: For previous month data
- `_get_first_of_current_month()`: For current month data

## 🐛 Troubleshooting

### Common Issues

1. **ChromeDriver version mismatch**
   - Ensure ChromeDriver matches your Chrome browser version

2. **Download timeout**
   - Increase timeout in `_wait_for_download()` method

3. **Element not found**
   - Update XPath/CSS selectors if website structure changes

4. **VPN connection issues**
   - Verify VPN client path and shortcut ID
   - Set `use_vpn=False` to disable VPN

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🤝 Contributing
Contributions are welcome! If you have suggestions or improvements, please open an issue or submit a pull request. Please follow standard coding conventions and include appropriate tests if applicable.

## 📧 Contact
For questions or feedback, feel free to reach out:

<p align="center"> <strong>📬 Let’s Connect!</strong> <br> <a href="https://www.linkedin.com/in/rishikesh-borah-3b245284/" target="_blank" style="display: inline-block; background-color: #0077B5; color: #fff; padding: 8px 16px; margin: 5px 10px; text-decoration: none; border-radius: 4px;">LinkedIn</a> | <a href="mailto:rishikesh.borah4@gmail.com" target="_blank" style="display: inline-block; background-color: #D44638; color: #fff; padding: 8px 16px; margin: 5px 10px; text-decoration: none; border-radius: 4px;">Email</a> </p>
---

_Stay curious, automate wisely! ✨_

# Fee Commission Selenium Automation

This script automates the fee-commission workflow shown in the screenshots:

1. Log in with credentials from `credentials.txt` or `credentials`
2. Open the `Add For Biller Merchant` page
3. Search and select the merchant by mobile/account number
4. Select the service and payee
5. Submit the header form
6. Enter the time
7. Add fee and commission rows from the Excel workbook for `UDDOKTA` and `CUSTOMER`
8. Register each completed payee configuration

## Files

- `main.py`: Selenium automation script
- `requirements.txt`: Python dependencies
- `credentials.txt.example`: sample credentials file format

## Expected workbook format

The script reads the first worksheet and looks for rows shaped like the screenshots:

- channel cells such as `APP` and `USSD`
- payee labels such as `UDDOKTA` and `CUSTOMER`
- player headers like `UDD`, `DH`, `MD`, `TWTL`, `BPO`, `AD`
- a `Service Fee` column

The default player mapping is:

- `UDD` -> `UDDOKTA`
- `DH` -> `DISTRIBUTOR`
- `AD` -> `ADVANCE COMMISSION`

## Setup

```powershell
python -m pip install -r requirements.txt
Copy-Item .\credentials.txt.example .\credentials.txt
```

Then fill in `credentials.txt`, or rename it to `credentials` if you want to match the screenshot naming.

Place the Excel workbook in this folder, or pass its location with `--workbook`.

## Run

```powershell
python .\main.py `
  --merchant-account-no 01901009906 `
  --service-query "asfia | 1333" `
  --effective-time-hh24 17
```

Optional flags:

- `--browser chrome`
- `--headless`
- `--workbook "D:\path\to\Fee-Com File Automation.xlsx"`
- `--credentials "D:\path\to\credentials.txt"`
- `--login-url "https://..."`
- `--system-url "https://.../ui/system/#/home"`

## Notes

- The portal locators are based on the screenshot flow and visible labels, so you may need small XPath adjustments if the live DOM differs.
- The script expects Selenium Manager to provision the browser driver automatically once Selenium is installed.

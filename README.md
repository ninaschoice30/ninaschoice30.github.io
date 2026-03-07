# Excel Macro: Travel Order Automation

This Excel workbook automates the creation and emailing of travel orders using a VBA macro. It copies a travel order sheet, populates it with user input, generates a temporary Excel file, and sends it via Outlook. After sending, the temporary file is automatically deleted for security.

---

## 📂 Files in this repository
- `travel_order_macro.xlsm` – Macro-enabled workbook with full functionality.  
---

## ⚙️ Features
- Automatically collects user inputs:  
  - Full Name  
  - Position  
  - Destination  
  - Travel start & end date/time  
  - Purpose of travel  
  - Duration of stay  
- Generates a new travel order sheet populated with the inputs.  
- Saves a temporary workbook copy for emailing.  
- Automatically opens Outlook and attaches the travel order for emailing.  
- Automatically deletes the temporary file after sending.  

---

## 📝 How to Use
1. Download `TravelOrderAutomation.xlsm`.  
2. Open the workbook in Excel and **enable macros**.  
3. Run the macro `sendemail` (Developer → Macros → Run).  
4. Follow the input prompts to provide travel details.  
5. Outlook will open a draft email with the travel order attached; review and send.  

---

## 🔗 Share & Connect
If you find this macro useful, feel free to share on LinkedIn or tag me!  

#ExcelVBA #Automation #Macros #TravelOrders #Productivity

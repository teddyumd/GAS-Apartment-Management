GAS-Apartment-Management
Google Apps Script-based Property Management System
Apartment Management System  

Description  
This project is a Google Apps Script-based Apartment Management System that helps landlords manage tenants, rent payments, and maintenance requests. It includes:  

- Code.gs → Handles the backend logic in Google Apps Script.  
- index.html → Provides the user interface using HTML, JavaScript, and Bootstrap.  

Features  
Tenant Management: Register new tenants and track lease details.  
Rent Payment Tracking: Log rent payments and automatically calculate the next due date.  
Maintenance Requests: Submit and manage maintenance requests for individual units.  
Automated Rent Payment ID: Generates and increments unique rent payment IDs.  
Pagination & Sorting: Ensures smooth navigation through large datasets.  

---

Installation & Deployment  

Step 1: Download this Excel file to your computer.
Upload it to Google Drive and open it with Google Sheets. Or you can access the file from here:
https://docs.google.com/spreadsheets/d/1j1U0ze76-tS5D0AA7PnEzCe2bj-QBNKQA54LxeDpEVw/edit?usp=sharing

Step 2: Convert to Google Sheets Format (if needed)
Click File > Save as Google Sheets to ensure compatibility with the script.

Step 3: Deploy the Google Apps Script (Follow the instructions in the README file)
- Open Google Apps Script
1. Go to Google Drive and create a new Google Apps Script project.  
2. Delete any default code in `Code.gs`.  

Step 4: Add Code.gs
1. Copy the contents of `Code.gs` and paste it into your Apps Script editor.  

Step 5: Add index.html
1. In the Apps Script project, click `+` to create a new file.  
2. Name it `index.html` and paste the contents from your `index.html` file.  

Step 6: Deploy the Web App
1. Click Deploy > New Deployment.  
2. Under Select type, choose `Web App`.  
3. Set Who has access to `Anyone`.  
4. Click Deploy and authorize the script.  
5. Copy the Web App URL and open it in a browser.  

---

License  
This project is licensed under the MIT License. See `LICENSE.txt` for details.  

---
What's in the Excel file for an Apartment Management System
- Tenants, Rent Payments, Maintenance Requests, Building Maintenance, Access List, and Units.
- Tenants sheet. The columns include Unit Number, Tenant Name, Phone Number, Email, Lease dates, Security Deposit, Status, Lease Amount, and Payment Frequency. The payment frequencies vary, such as Monthly, Every 2 Months.
- Rent Payments sheet has details like Rent Payment ID, Unit Number, Tenant Name, Frequency, Payment Date, Amount, Next Due Date, Status, Notes, and a Payment URL Link. The URLs will be used for payment receipts stored in Google Drive.
- Maintenance Requests include Request ID, Unit Number, Tenant Name, Issue Description, Date Submitted, Status, Date Resolved, Days to Complete, and Notes.
- Building Maintenance has Maintenance REQ ID, Task Name, Due Date, Completion Status, Expense Amount, Date Completed, and Notes.
- Access List sheet includes User Emails and Roles (Tenant/Property Manager). This is used for security purpose. The web application will check to see if users have access based on this list. The application is configured to provide access to only the property manager role. If anyone wants to change this to Tenants, this is possible through the code.
- Units sheet lists Unit Numbers and Status (Rented/Vacant). This is used to update the modals and cards on the web application.
---
Credits  
Developed by Tewodros Hailegeberel.  

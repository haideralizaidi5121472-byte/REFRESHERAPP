Refresher H2o Water Delivery App

Features
Admin
Customer management with auto generated customer ids
Per customer water type and one time customer price
Orders view and rider assignment
Dashboard with today sales, today deliveries, outstanding, cash, jazzcash, expenses
Monthly charts for sales, recovery, expenses
Import customers from Excel or CSV and auto generate ids
Daily expenses entry

Customer
Login by customer id
See selected water type, list price, customer price, discount per bottle
Place order with quantity
Tracking with totals and recent orders
Outstanding balance

Rider
Login by rider id
See assigned orders automatically
Update status accepted and on the way
Complete order with delivered bottles, empty returned, payment type and amount

How to run locally
1 Install Node.js 18 or later
2 In the project folder run
  npm install
  npm start
3 Open
  http://localhost:3000

Default admin
Username admin
Password refresher123

Change admin password before first start
Set environment variable ADMIN_PASSWORD

Example
Windows PowerShell
  $env:ADMIN_PASSWORD="MyNewPass"
  npm start

Linux or Mac
  ADMIN_PASSWORD="MyNewPass" npm start
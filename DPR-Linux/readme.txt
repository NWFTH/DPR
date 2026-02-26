If you just want to run a one-time sync for a specific month in 2025, 
change these lines at the top of your runRpa function:

// Change this:
const currentYear = now.getFullYear(); 

// To this:
const currentYear = 2025; 
const currentMonth = "Dec"; // Or whichever month you need from 2025

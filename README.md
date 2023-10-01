# Invoice-Generator-and-Manager
Created a Tkinter-driven software solution for seamless invoice generation, storage, printing, and pinpoint retrieval by specific dates.
User-Friendly Interface: 
Craft an intuitive UI with Tkinter widgets, ensuring effortless navigation and task execution.

## Effortless Invoice Creation: 
Allows users to effortlessly craft invoices, inputting customer details, items, and quantities. Automatic total calculation and unique invoice IDs streamline the process.

## Database Integration: 
Utilize an integrated MongoDb database for secure and swift data storage, ensuring easy access to historical invoices.

## Date-Centric Search: 
Implemented a robust search and filter feature enabling users to swiftly locate invoices generated on particular dates.

## Organized Invoice Display: 
It presents invoices in a well-organized list, enabling scrolling, sorting, and searches by various criteria.

## Seamless Printing: 
Empower users to print invoices directly from the application, leveraging Tkinter's printing capabilities.

## Flexible Saving Options: 
Offers versatile invoice saving formats, including PDF, for effortless sharing and archiving.

## Error-Resilient: 
Incorporated rigorous error handling to guide users through input and database challenges.

## Thorough Testing: 
Ensured software reliability through comprehensive testing, enhancing user satisfaction.

## Effortless Distribution: 
Package the software as an executable file for easy distribution, eliminating the need for separate Python or Tkinter installations.

With this Tkinter-driven invoice management solution, users can effortlessly handle invoice tasks, from creation and saving to printing and precise retrieval, all within an elegantly designed interface.

# Setup
## Step 1:
Set all the paths to your respective folder path the code 
### The code automatically takes your base path
## step 2:
connect to database (In my case it is MongoDB)
## step 3:
now your code is ready to be converted to .exe 
```
pip install pyinstaller
```
```
pyinstaller --name Invoice_Generated --onefile --windowed --icon=icon.ico main.py
```

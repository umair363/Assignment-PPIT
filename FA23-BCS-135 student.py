from tkinter import *
import pyodbc
6
try:
    conn = pyodbc.connect(
        r'Driver={Microsoft Access Driver (*.mdb, *.accdb)};DBQ=E:\ICT Assignment\Database41.accdb;')
    cursor = conn.cursor()
    print("Connected To Database")
except pyodbc.Error as e:
    print("Error in Connection", e)


def clear_record():
    for entry in (SnameEntery, FnameEntery, CnicEntery, CityEntery, marksEntery):
        entry.delete(0, END)
        message.config(text="Record Cleared ", fg="green")
        print("Record Cleared")


def FirstRecord():
    cursor.execute('SELECT Sname, Fname, Cnic, City, marks FROM Student ORDER BY Cnic ASC')
    row = cursor.fetchone()
    if row:
        display_record_in_text_boxes(row)
        message.config(text="First Record of Student Table", fg="green")
    else:
        message.config(text="No records available", fg="red")


# NextRecord function
def NextRecord():
    current_cnic = CnicValue.get()
    cursor.execute('SELECT * FROM Student WHERE Cnic > ? ORDER BY Cnic ASC', (current_cnic,))
    row = cursor.fetchone()
    if row:
        display_record_in_text_boxes(row)
        message.config(text="Next Record of Student Table", fg="green")

    else:
        message.config(text="No more records available", fg="red")


# PreviousRecord function
def PreviousRecord():
    current_cnic = CnicValue.get()
    cursor.execute('SELECT * FROM Student WHERE Cnic < ? ORDER BY Cnic DESC', (current_cnic,))
    row = cursor.fetchone()
    if row:
        display_record_in_text_boxes(row)
        message.config(text="Previous Record of Student Table", fg="green")
    else:
        message.config(text="No previous records available", fg="red")


def LastRecord():
    cursor.execute('SELECT Sname, Fname, Cnic, City, marks FROM Student ORDER BY Cnic DESC')
    row = cursor.fetchone()
    if row:
        display_record_in_text_boxes(row)
        message.config(text="Last Record of Student Table", fg="green")
    else:
        message.config(text="No records available", fg="red")


# InsertRecord function
def InsertRecord():
    name = SnameValue.get()
    FName = FnameValue.get()
    Cnic = CnicValue.get()
    City = CityValue.get()
    marks = marksValue.get()

    cursor.execute("SELECT Cnic FROM Student WHERE Cnic = ?", (Cnic,))
    existing_cnic = cursor.fetchone()

    if existing_cnic:
        message.config(text="CNIC Already Saved!Enter Other!", fg="red")
        print("CNIC already exists! Try another.")
    else:
        cursor.execute("INSERT INTO Student (Sname, Fname, Cnic, City, marks) VALUES (?, ?, ?, ?, ?)",
                       (name, FName, Cnic, City, marks))
        conn.commit()
        message.config(text="Record inserted successfully", fg="green")
        print("Insert Record")


def UpdateRecord():
    name = SnameValue.get()
    FName = FnameValue.get()
    Cnic = CnicValue.get()
    City = CityValue.get()
    marks = marksValue.get()

    # Check if CNIC field is empty
    if not Cnic.strip():  # Using strip() to remove leading/trailing whitespaces
        message.config(text="Please Enter CNIC# First!", fg="red")
        print("Please Enter CNIC# First!")
        return

    # Retrieve the existing record
    cursor.execute("SELECT * FROM Student WHERE Cnic = ?", (Cnic,))
    existing_record = cursor.fetchone()

    if not existing_record:
        message.config(text="Please enter correct CNIC to update!", fg="red")
        print("Please enter correct CNIC to update!")
        return

    # Check if data has changed except for CNIC
    existing_data = existing_record[0:4]  # Considering first 4 fields excluding CNIC
    new_data = (name, FName, City, marks)

    if existing_data == new_data:
        message.config(text="No changes detected to update!", fg="red")
        print("No changes detected to update!")
        return

    # Perform update if data is modified
    cursor.execute("UPDATE Student SET Sname=?, Fname=?, City=?, marks=? WHERE Cnic=?",
                   (name, FName, City, marks, Cnic))
    conn.commit()
    message.config(text="Record updated successfully", fg="green")
    print("Update Record")


# DeleteRecord function
def DeleteRecord():
    cnic = CnicValue.get()

    # Check if CNIC field is empty
    if not cnic.strip():  # Using strip() to remove leading/trailing whitespaces
        message.config(text="Please Enter CNIC# First!", fg="red")
        print("Please Enter CNIC# First!")
        return

    # Check if the CNIC exists in the database
    cursor.execute("SELECT Cnic FROM Student WHERE Cnic = ?", (cnic,))
    existing_cnic = cursor.fetchone()

    if not existing_cnic:
        message.config(text="No Record Found to Delete!", fg="red")
        print("No Record Found to Delete!")
        return

    # Proceed with the deletion if CNIC exists
    cursor.execute("DELETE FROM Student WHERE Cnic=?", (cnic,))
    conn.commit()
    message.config(text="Record deleted successfully", fg="green")
    print("Delete Record")

    # Proceed with the deletion if CNIC is provided
    cursor.execute("DELETE FROM Student WHERE Cnic=?", (cnic,))
    conn.commit()
    message.config(text="Record deleted successfully", fg="green")
    print("Delete Record")


# SearchRecord function

def SearchRecord():
    search_term = CnicValue.get()

    cursor.execute("SELECT * FROM Student WHERE Cnic LIKE ?", ('%' + search_term + '%',))
    row = cursor.fetchone()

    if row:
        display_record_in_text_boxes(row)
        message.config(text="Record searched successfully", fg="green", font=("Ariel", 15, "bold"))
    else:
        message.config(text="No records found", fg="red", font=("Ariel", 15, "bold"))
    print("Item Searched Sucessfully!")


# to display records on student form
def display_record_in_text_boxes(row):
    SnameValue.set(row[0])
    FnameValue.set(row[1])
    CnicValue.set(row[2])
    CityValue.set(row[3])
    marksValue.set(row[4])


# Design the Student Database Form
root = Tk()
root.geometry("600x400")

SnameValue = StringVar()
FnameValue = StringVar()
CnicValue = StringVar()
CityValue = StringVar()
marksValue = DoubleVar()

SnameEntery = Entry(root, textvariable=SnameValue, width='30', font='ar 12 bold')
FnameEntery = Entry(root, textvariable=FnameValue, width='30', font='ar 12 bold')
CnicEntery = Entry(root, textvariable=CnicValue, width='30', font='ar 12 bold')
CityEntery = Entry(root, textvariable=CityValue, width='30', font='ar 12 bold')
marksEntery = Entry(root, textvariable=marksValue, width='30', font='ar 12 bold')

Label(root, text="Student Database Form", font="Arial 12 bold", foreground='blue').grid(row=0, column=0)
message = Label(root, text="Message Will Appear Here!", foreground='red')
message.grid(row=0, column=1)
Sname = Label(root, text='Student Name', font="ar 10 bold")
Fname = Label(root, text='Father Name', font="ar 10 bold")
Cnic = Label(root, text='Cnic# (P.Key)', font="ar 10 bold")
search = Label(root, text='Search Record', font="ar 10 bold")
City = Label(root, text='City', font="ar 10 bold")
marks = Label(root, text='Marks', font="ar 10 bold")

Sname.grid(row=2, column=0)
Fname.grid(row=3, column=0)
Cnic.grid(row=4, column=0)
search.grid(row=4, column=2)
City.grid(row=5, column=0)
marks.grid(row=6, column=0)

SnameEntery.grid(row=2, column=1, pady=15)
FnameEntery.grid(row=3, column=1, pady=15)
CnicEntery.grid(row=4, column=1, pady=15)
CityEntery.grid(row=5, column=1, pady=15)
marksEntery.grid(row=6, column=1, pady=15)
Button(text="CLEAR", command=clear_record, background='gray', foreground='blue', font='ar 10 bold').grid(row=7,
                                                                                                         column=0)
Button(text="FIRST", command=FirstRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=7, column=1)
Button(text="NEXT", command=NextRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=7, column=2)
Button(text="PREVIOUS", command=PreviousRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=9,
                                                                                                              column=0)
Button(text="LAST", command=LastRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=9, column=1)
Button(text="INSERT", command=InsertRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=9,
                                                                                                          column=2)
Button(text="UPDATE", command=UpdateRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=11,
                                                                                                          column=0)
Button(text="DELETE", command=DeleteRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=11,
                                                                                                          column=1)
Button(text="SEARCH", command=SearchRecord, background='gray', foreground='blue', font='ar 10 bold').grid(row=11,
                                                                                                          column=2)

root.mainloop()
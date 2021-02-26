from tkinter import *
from tkinter.ttk import *
from tkinter import scrolledtext
from tkinter import messagebox
import csv
import tkinter as tk
from string import *
import sqlite3
import webbrowser
from tksheet import *
w=Tk()
w.configure(background='steel blue')
w.title("PON VIDYASHRAM STUDENT DATBASE")
a=Label(w, text ='PON VIDYASHRAM STUDENT DATABASE', font = ('Helvetica',20)) 
a.pack()
w.geometry('1200x1300')
studentname=StringVar()
rollno=StringVar()
standard=StringVar()
section=StringVar()
fathername=StringVar()
phonenumber=StringVar()
emailid=StringVar()
address=StringVar()
aadharno=StringVar()
t1=Label(w,text='ENTER DETAILS',font =('Helvetica',25))
t1.place(x=513,y=345)
t2=Label(w,text='A PREVIEW OF EXCEL',font =('Helvetica',25))
t2.place(x=450,y=395)
def connection():
    try:
        conn=sqlite3.connect("student database.db")
    except:
        t1.configure(text="CANNOT CONNECT TO DATABASE\n")
    return conn

def verifier():
    a=b=c=d=e=f=g=h=i=0
    if not studentname.get():
        t1.configure(text="ROLL NO IS REQUIRED\n")
        a=1
    if not rollno.get():
        t1.configure(text="STUDENT NAME IS REQUIRED\n")
        b=1
    if not standard.get():
        t1.configure(text="<>STANDARD IS REQUIRED\n")
        c=1
    if not section.get():
        t1.configure(text="SECTION IS REQUIRED\n")
        d=1
    if not fathername.get():
        t1.configure(text="FATHER'S NAME IS REQUIRED\n")
        e=1
    if not phonenumber.get():
        t1.configure(text="<>PHONE NUMBER IS REQUIRED\n")
        f=1
    if not emailid.get():
        t1.configure(text="EMAIL ID IS REQUIRED\n")
        g=1    
    if not e6.get('1.0', tk.END):
        t1.configure(text="ADRESS IS REQUIRED\n")
        h=1
    if not aadharno.get():
        t1.configure(text="AADHAR NO IS REQUIRED\n")
        i=1    
    if a==1 or b==1 or c==1 or d==1 or e==1 or f==1 or g==1 or h==1 or i==1:
        return 1
    else:
        return 0                
    


def add_student():
            ret=verifier()
            if ret==0:
                conn=connection()
                cur=conn.cursor()
                d=(e1.get(),e2.get(),standard.get(),section.get(),e3.get(),e4.get(),e5.get(),e6.get('1.0', 'end-1c'),e7.get())
                cur.execute('CREATE TABLE IF NOT EXISTS student (rollno integer,studentname text,standard integer, section text, fathername text,phonenumber integer,emailid text,address text,aadharno integer)')
                cur.execute("insert into student values(?,?,?,?,?,?,?,?,?)",d)
                conn.commit()
                conn.close()
            view_student()    
            e1.delete(0,END)
            e2.delete(0,END)
            e3.delete(0,END)
            e4.delete(0,END)
            e5.delete(0,END)
            e6.delete('1.0',END)
            e7.delete(0,END)
            standard.delete(0,END)
            section.delete(0,END)        
def addcsv():
    conn=connection()
    cur=conn.cursor()
    cur.execute("select * from student order by rollno")
    data=cur.fetchall()
    conn.close()
    l=['ROLL NO:','STUDENT NAME :','STANDARD:','SECTION:',"FATHER'SNAME:",'PHONE NUMBER:','EMAIL ID:','ADDRESS:','AADHAR NO.:']
    with open ('student database.csv',mode='w')as csvfile:
         mywriter=(csv.writer(csvfile))
         mywriter.writerow(l)
         for k in data:
             mywriter.writerow(k)
         
    webbrowser.open_new('student database.csv')
    
def view_student():
    conn=connection()
    cur=conn.cursor()
    cur.execute("select * from student order by rollno")
    data=cur.fetchall()
    conn.close()
    l=['ROLL NO:','STUDENT NAME :','STANDARD:',"FATHER'SNAME:",'PHONE NUMBER:','EMAIL ID:','AADHAR NO.:','ADDRESS:']
    sheet =Sheet(w,width=1200,height=400,row_height = "4")
    sheet.place(x=15,y=450)
    sheet.set_sheet_data([[f"{cj}" for cj in ri] for ri in data])        
    

      
def deletestudent():
    ede=int(ed.get())
    if not ed.get():
        t1.insert(END,"<>ROLL NO IS REQUIRED<>\n")
    else:
        conn=connection()
        cur=conn.cursor()
        a=int(ed.get())
        cur.execute("Delete from student where rollno = "+(ed.get()))
        cur.execute("select * from student order by rollno")
        data=cur.fetchall()
        cur.execute("update student set rollno=rollno-1 where rollno>"+(ed.get()))               
        conn.commit()
        conn.close()
        t1.configure(text="SUCCESSFULLY DELETED THE USER\n")        
        ed.delete(0,END)
        view_student()

def update_student():
    conn=connection()
    cur=conn.cursor()
    if e1.get():
               cur.execute("UPDATE student SET rollno=? where rollno=?",(e1.get(),eu.get()))
    if e2.get():
               cur.execute("UPDATE student SET studentname=? where rollno=?",(e2.get(),eu.get()))
    if standard.get():
               cur.execute("UPDATE student SET standard=? where rollno=?",(standard.get(),eu.get()))
    if section.get():
               cur.execute("UPDATE student SET section=? where rollno=?",(section.get(),eu.get()))
    if e3.get():
               cur.execute("UPDATE student SET fathername=? where rollno=?",(e3.get(),eu.get()))
    if e4.get():
               cur.execute("UPDATE student SET phonenumber=? where rollno=?",(e4.get(),eu.get()))
    if e5.get():
               cur.execute("UPDATE student SET emailid=? where rollno=?",(e5.get(),eu.get()))
    if e6.get('1.0', tk.END):
               cur.execute("UPDATE student SET address=? where rollno=?",(e6.get('1.0', 'end-1c'),eu.get()))
    if e7.get():
               cur.execute("UPDATE student SET aadharno=? where rollno=?",(e7.get(),eu.get()))
    conn.commit()
    conn.close()
    view_student()     
    e1.delete(0,END)
    e2.delete(0,END)
    e3.delete(0,END)
    e4.delete(0,END)
    e5.delete(0,END)
    e6.delete('1.0',END)
    e7.delete(0,END)
    standard.delete(0,END)
    section.delete(0,END)
    eu.delete(0,END)
    view_student()
def close():
    messagebox.showinfo('STUDENT DATABASE', 'ARE YOU SURE TO QUIT THE PROGRAM')
    w.destroy()
try:
    view_student()
except:
        t1.configure(text="CREATE DATA BASE\n")
    
l=['ROLL NO:','STUDENT NAME :','STANDARD:',"FATHER'SNAME:",'PHONE NUMBER:','EMAIL ID:','AADHAR NO.:','ADDRESS:']
for i in range(len(l)):
    a=Label(w,text=l[i],font =('Helvetica',10))
    a.place(x=15,y=(i*40)+40)
data=[]    
standard= Combobox(w,width=10)
standard['values']= (1,2,3,4,5,6,7,8,9,10,11,12)
standard.place(x=140,y=120)

sectionl=Label(w,text='SECTION:',font =('Helvetica',10))
sectionl.place(x=230,y=120)
section = Combobox(w,width=10)
section['values']= ('A','B','C','D','E','F','G')
section.place(x=310,y=120)

e1=Entry(w,textvariable=rollno,width=10)
e1.place(x=140,y=40)

e2=Entry(w,textvariable=studentname,width=40)
e2.place(x=140,y=80)

e3=Entry(w,textvariable=fathername,width=40)
e3.place(x=140,y=160)

e4=Entry(w,textvariable=phonenumber,width=40)
e4.place(x=140,y=200)
    
e5=Entry(w,textvariable=emailid,width=40)
e5.place(x=140,y=240)

e6= scrolledtext.ScrolledText(w,width=40,height=4)
e6.place(x=140,y=320) 

e7=Entry(w,textvariable=aadharno,width=40,)
e7.place(x=140,y=280)

ed=Entry(w,textvariable=deletestudent,width=10,)
ed.place(x=800,y=45)

eu=Entry(w,textvariable=update_student,width=10)
eu.place(x=800,y=105)
lu=Label(w,text='ENTER THE ROLL NO TO UPDATE STUDENT',font =('Helvetica',10))
lu.place(x=510,y=135)

ld=Label(w,text='ENTER THE ROLL NO TO DELETE STUDENT',font =('Helvetica',10))
ld.place(x=510,y=75)
b1=Button(w,text="DELETE STUDENT",command=deletestudent,width=40)
b1.place(x=510,y=45)

b2=Button(w,text="VIEW ALL STUDENTS",command=view_student,width=40)
b2.place(x=510,y=275)

b3=Button(w,text="UPDATE INFO",command=update_student,width=40)
b3.place(x=510,y=105)

b4=Button(w,text="CLOSE",width=40,command=close)
b4.place(x=510,y=315)

b5=Button(w,text="SAVE AS CSV EXCEL",width=40,command=addcsv)
b5.place(x=510,y=235)

b6=Button(w,text=" ADD STUDENT",width=40,command=add_student )
b6.place(x=510,y=175)

w.mainloop()

        
            
           


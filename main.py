from tkinter import *
from tkinter import messagebox,filedialog
from pygame import mixer
import speech_recognition
from email.message import EmailMessage
import smtplib
import os
import imghdr
import pandas
check=False



def iexit():
    result=messagebox.askyesno("Notification", "Do you want to exit?")
    if result:
       root.destroy()
    else:
        pass
    
def clear():
    toEntryField.delete(0,END)
    subjectEntryField.delete(0,END)
    textarea.delete(1.0,END)

def speak():
    mixer.init()
    mixer.music.load("music1.mp3")
    mixer.music.play()
    sr = speech_recognition.Recognizer()
    
    with speech_recognition.Microphone() as m:
        try:
            sr.adjust_for_ambient_noise(m, duration=0.2)
            print("Listening...")
            audio = sr.listen(m)
            print("Recognizing...")
            try:
                text = sr.recognize_google(audio)
                textarea.insert(END, text + ".\n")
                print("You said: " + text)
            except speech_recognition.UnknownValueError:
                messagebox.showerror("Error", "Could not understand audio", parent=root)
            except speech_recognition.RequestError as e:
                messagebox.showerror("Error", f"Could not request results; {e}", parent=root)
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=root)

def settings():
    def clear1():
        fromEntryField.delete(0,END)
        paswEntryField.delete(0,END)

    def save():
       if fromEntryField.get()=="" or paswEntryField.get()=="":
            messagebox.showerror('Error',"All Fields are Required",parent=root1)
       else:
           f=open("credentials.txt",'w')
           f.write(fromEntryField.get()+","+paswEntryField.get())
           f.close()
           messagebox.showinfo("Information","Credentials Saved Successfully",parent=root1)

       
        
    root1=Toplevel()
    root1.title('Setting')
    root1.geometry("650x340+550+90")
    root1.config(bg="dodger blue2")

    Label(root1,text="Credential Settings",image=logoImage,compound=LEFT,font=('Goudy Old Style',40,'bold'),fg="white",bg="gray20").grid(padx=60)

    fromLabelFrame=LabelFrame(root1,text='from (Email Address)',font=('times new roman',16,'bold'),bd=5,fg='white',bg='dodger blue2')
    fromLabelFrame.grid(row=1,column=0,pady=20)

    fromEntryField=Entry(fromLabelFrame,font=('times new roman',18,'bold'),width=30)
    fromEntryField.grid(row=0,column=0)

    paswLabelFrame=LabelFrame(root1,text='Password',font=('times new roman',16,'bold'),bd=5,fg='white',bg='dodger blue2')
    paswLabelFrame.grid(row=2,column=0,pady=20)

    paswEntryField=Entry(paswLabelFrame,font=('times new roman',18,'bold'),width=30,show="*")
    paswEntryField.grid(row=0,column=0)

    Button(root1,text="Save",font=('times new roman',18,'bold'),cursor='hand2',bg="gold2",fg="black",command=save).place(x=210,y=280)
    Button(root1,text="Clear",font=('times new roman',18,'bold'),cursor='hand2',bg="gold2",fg="black",command=clear1).place(x=340,y=280)

    f=open('credentials.txt','r')
    for i in f:
       credentials=i.split(",")
    fromEntryField.insert(0,credentials[0])
    paswEntryField.insert(0,credentials[1])

    root1.mainloop()


def sendingEmail(toaddress,subject,body):
    f=open("credentials.txt",'r')
    for i in f:
       credentials=i.split(",")

    message=EmailMessage()
    message['subject']=subject
    message['to']=toaddress
    message['from']=credentials[0]
    message.set_content(body)

    if check:

        if file_type=="png" or file_type=="jpeg" or file_type=="jpg":
            f=open(file_path,'rb')
            file_data=f.read()
            subtype=imghdr.what(file_path)
            message.add_attachment(file_data,maintype='image',subtype=subtype,filename=file_name)
        else:
            f=open(file_path,'rb')
            file_data=f.read()

            message.add_attachment(file_data,maintype="application",subtype="octet-stream",filename=file_name)

    s=smtplib.SMTP('smtp.gmail.com',587)
    s.starttls()
    s.login(credentials[0],credentials[1])
    s.send_message(message)

    x=s.ehlo()
    if x[0]==250:
        return "sent"
    else:
        return "Failed"
    



def send():
    
    if toEntryField.get()=="" or subjectEntryField.get()=="" or textarea.get(1.0,END)=="\n":
       messagebox.showerror('Error',"Fields are Required",parent=root)
    else:
        
        if choice.get()=="single":
            result=sendingEmail(toEntryField.get(),subjectEntryField.get(),textarea.get(1.0,END))

            if result=="sent":
                messagebox.showinfo("Information","Message Sent")

            if result == "failed":
                messagebox.showerror("Error","Email is not Sent")

        if choice.get()=="multiple":
            sent=0
            failed=0
            for x in final_emails:
                result=sendingEmail(x,subjectEntryField.get(),textarea.get(1.0,END))
                if result=="sent":
                    sent+=1
                if result=="Failed":
                    failed+=1

                totalLabel.config(text="")
                sentLabel.config(text="Sent:" + str(sent))
                leftLabel.config(text="Left:" + str(len(final_emails)-(sent-failed)))
                failedLabel.config(text="Failed:" + str(failed))

                totalLabel.update()
                sentLabel.update()
                leftLabel.update()
                failedLabel.update()
            messagebox.showinfo("Success","Mails sent Successfully")

        



    
def attachment():
    global file_name,file_type,file_path,check
    check=True
    file_path=filedialog.askopenfilename(initialdir='c:/',title="Select File")
    file_type=file_path.split(".")
    file_type=file_type[1]
    file_name=os.path.basename(file_path)
    textarea.insert(END,f"\n{file_name}\n")


def button_check():
    if choice.get()=="multiple":
        browseButton.config(state=NORMAL)
        toEntryField.config(state="readonly")
    if choice.get()=="single":
        browseButton.config(state=DISABLED)
        toEntryField.config(state="readonly")

def browse():
    global final_emails
    path=filedialog.askopenfilename(initialdir="c:/",title="Select Excel File")
    if path=="":
        messagebox.showerror("Error","Please select an excel file")
    else:
        data=pandas.read_excel(path)

        if "Email" in data.columns:
            emails=list(data['Email'])
            final_emails=[]

            for i in emails:
                if pandas.isnull(i)==False:
                    final_emails.append(i)
            if len(final_emails)==0:
                messagebox.showerror("Error","File doesn't contain any email addresses")
            else:
                toEntryField.config(state=NORMAL)
                toEntryField.insert(0,os.path.basename(path))
                totalLabel.config(text="Total: "+str(len(final_emails)))
                sentLabel.config(text="Sent: ")
                leftLabel.config(text="Left:")
                failedLabel.config(text="Failed:")
           
root = Tk()
root.title("Email sender app")
root.geometry('780x620+100+50')
root.resizable(0,0)
root.config(bg="dodger blue2")


#setting icon images
titleFrame=Frame(root, bg="white")
titleFrame.grid(row=0,column=0)
logoImage=PhotoImage(file='email.png')
titleLable=Label(titleFrame,text="      Email Sender",image=logoImage,compound=LEFT,font=('Goudy Old Style',28,'bold'),bg='white')
titleLable.grid(row=0,column=0)
settingImage=PhotoImage(file='setting.png')
Button(titleFrame,image=settingImage,bd=0,bg='white',cursor="hand2",activebackground='white',command=settings).grid(row=0,column=1,padx=40)


#setting single and multiple radio buttons
chooseFrame=Frame(root,bg="dodger blue2")
chooseFrame.grid(row=1,column=0,padx=20)
choice=StringVar()
sbutton=Radiobutton(chooseFrame,text='Single',font=("times new roman",25),variable=choice,value='single',bg="dodger blue2",activebackground='dodger blue2',command=button_check)
sbutton.grid(row=0,column=0,padx=20)
mbutton=Radiobutton(chooseFrame,text='Multiple',font=("times new roman",25),variable=choice,value='multiple',bg="dodger blue2",activebackground='dodger blue2',command=button_check)
mbutton.grid(row=0,column=1)

choice.set('single')

toLabelFrame=LabelFrame(root,text='To (Email Address)',font=('times new roman',16,'bold'),bd=5,fg='white',bg='dodger blue2')
toLabelFrame.grid(row=2,column=0)

toEntryField=Entry(toLabelFrame,font=('times new roman',18,'bold'),width=30)
toEntryField.grid(row=0,column=0)

browseImage=PhotoImage(file='browse.png')

browseButton=Button(toLabelFrame,text=' Browse',image=browseImage,compound=LEFT,font=('arial',12,'bold'),
       cursor='hand2',bd=0,bg='dodger blue2',activebackground='dodger blue2',state=DISABLED,command=browse)
browseButton.grid(row=0,column=1,padx=20)

subjectLabelFrame=LabelFrame(root,text='Subject',font=('times new roman',16,'bold'),bd=5,fg='white',bg='dodger blue2')
subjectLabelFrame.grid(row=3,column=0,pady=10)

subjectEntryField=Entry(subjectLabelFrame,font=('times new roman',18,'bold'),width=30)
subjectEntryField.grid(row=0,column=0)

emailLabelFrame=LabelFrame(root,text='Compose Email',font=('times new roman',16,'bold'),bd=5,fg='white',bg='dodger blue2')
emailLabelFrame.grid(row=4,column=0,padx=20)
micImage=PhotoImage(file='mic.png')


Button(emailLabelFrame,text=' Speak',image=micImage,compound=LEFT,font=('arial',12,'bold'),
       cursor='hand2',bd=0,fg='black',bg='dodger blue2',activebackground='dodger blue2',command=speak).grid(row=0,column=0)
attachImage=PhotoImage(file='attachments.png')

Button(emailLabelFrame,text=' Attachment',image=attachImage,compound=LEFT,font=('arial',12,'bold'),
       cursor='hand2',bd=0,bg='dodger blue2',activebackground='dodger blue2',command=attachment).grid(row=0,column=1)


textarea=Text(emailLabelFrame,font=('times new roman',14,),height=8)
textarea.grid(row=1,column=0,columnspan=2)


sendImage=PhotoImage(file='send.png')
Button(root,image=sendImage,bd=0,bg='dodger blue2',cursor='hand2',activebackground='dodger blue2',command=send).place(x=490,y=540)


clearImage=PhotoImage(file='clear.png')
Button(root,image=clearImage,bd=0,bg='dodger blue2',cursor='hand2',activebackground='dodger blue2',command=clear).place(x=590,y=550)


exitImage=PhotoImage(file='exit.png')
Button(root,image=exitImage,bd=0,bg='dodger blue2',cursor='hand2',activebackground='dodger blue2',command=iexit).place(x=690,y=550)


totalLabel=Label(root,font=('times new roman',18,'bold'),bg='dodger blue2',fg='black')
totalLabel.place(x=10,y=560)


sentLabel=Label(root,font=('times new roman',18,'bold'),bg='dodger blue2',fg='black')
sentLabel.place(x=100,y=560)


leftLabel=Label(root,font=('times new roman',18,'bold'),bg='dodger blue2',fg='black')
leftLabel.place(x=190,y=560)


failedLabel=Label(root,font=('times new roman',18,'bold'),bg='dodger blue2',fg='black')
failedLabel.place(x=280,y=560)


root.mainloop()
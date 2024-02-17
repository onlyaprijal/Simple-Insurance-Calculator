from tkinter import *
from tkcalendar import Calendar
from tkinter import messagebox,font,filedialog
import numpy as np
import pandas as pd
data = pd.read_excel('TMI_IV_versi_Excel.xlsx')


def is_number(value):
    try:
        float_number = float(value)
        return True
    except ValueError:
        return False

def show_calender(entrydate):
    top = Toplevel()
    calender = Calendar(top, font="Arial 14", selectmode='day', locale='en_US')
    calender.pack(fill="both", expand=True)

    if entrydate == entry_born:
        top.title('Insured Born Date')
    elif entrydate == entry_start:
        top.title('Insurance Start Date')

    def get_date():
        date = calender.selection_get()
        entrydate.delete(0,END)
        entrydate.insert(0, date.strftime('%d-%m-%Y'))
        top.destroy()

    button_get_date = Button(top,text='Enter Date',command=get_date,bg="#E0E0E0")
    button_get_date.pack() 


def get_insured_data():


    global root
    root = Tk()
    root.title('Simple Insurance Calculator')

    name_label = Label(root,text='Enter Insured Name : ')
    global name_input
    name_input = Entry(root,width=30,borderwidth=10)

    insured_born_date_label = Label(root,text='Enter Insured Born Date : ')
    global entry_born
    entry_born = Entry(root,width=30,borderwidth=10)
    entry_born.bind("<Button-1>", lambda show: show_calender(entry_born))

    insurance_start_date_label = Label(root,text='Enter Insurance Start Date : ')
    global entry_start
    entry_start = Entry(root, width=30,borderwidth=10)
    entry_start.bind("<Button-1>", lambda show: show_calender(entry_start))

    list_sex = ['male','female']
    sex_string = StringVar()
    sex_label = Label(root,text='Select The Insured Sex : ')
    sex_option = OptionMenu(root, sex_string, *list_sex)
    

    def data_validation():
        global name, birth_date, insurance_start_date, sex

        name = name_input.get()
        birth_date = entry_born.get()
        insurance_start_date = entry_start.get()
        sex = sex_string.get()

    
        response = messagebox.askyesno(
        title='Is your input already correct?',
        message=f'Insured Name\t\t: {name}\n'
                f'Insured Born Date\t\t: {birth_date}\n'
                f'Insurance Start Date\t: {insurance_start_date}\n'
                f'Insured sex\t\t: {sex}\n')
    
        if response == 1 :
            root.destroy()

        else :
            entry_born.delete(0,END)
            entry_start.delete(0,END)  

    button_get_data = Button(root,text='Submit',relief='raised',bg="#E0E0E0",command=data_validation)

    name_label.pack(fill=BOTH)
    name_input.pack(fill=BOTH)
    insured_born_date_label.pack(fill=BOTH)
    entry_born.pack(fill=BOTH)
    insurance_start_date_label.pack(fill=BOTH)
    entry_start.pack(fill=BOTH)
    sex_label.pack(fill=BOTH)
    sex_option.pack(fill=BOTH)
    button_get_data.pack(fill=BOTH)
    root.mainloop()

    return name, birth_date, insurance_start_date, sex

def update_entry_state(event,insurance_duration_entry,insurance_deferred_entry,second_benefit_entry):
    selected_product = product_str.get()

    if selected_product == 'Wholelife insurance':
        insurance_duration_entry.config(state='disabled')
    else:
        insurance_duration_entry.config(state='normal')

    if selected_product in ['Pure endowment insurance', 'Endowment insurance']:
        insurance_deferred_entry.config(state='disabled')    
    else:
        insurance_deferred_entry.config(state='normal')

    if selected_product == 'Endowment insurance':
        second_benefit_entry.config(state='normal')

    else :
        second_benefit_entry.config(state='disabled')

def insurance_options(insurance_start_date,birth_date):
    root = Tk()
    root.title('Simple Insurance Calculator')

    Label1 = Label(root, text='Select the Insurance Product',font=font.Font(size=12),anchor='center')
    insurance_product = [
        'Wholelife insurance', 'N-years term insurance',
        'Pure endowment insurance', 'Endowment insurance'
    ]

    Label2 = Label(root, text='Select your fractional age assumption',font=font.Font(size=12),anchor='center')

    global product_str
    product_str = StringVar()
    insurance_product_option = OptionMenu(root, product_str, *insurance_product, command=lambda selected_product=product_str: update_entry_state(selected_product, insurance_duration_entry, insurance_deferred_entry,second_benefit_entry))

    assumption_str = StringVar()
    assumption_list = ['UDD', 'CF', 'Hyperbolic']
    assumption_option = OptionMenu(root, assumption_str, *assumption_list)

    Label3 = Label(root, text='Insert the fractional insurance interval (annually)',font=font.Font(size=12),anchor='center')
    fractional_insurance_interval_entry = Entry(root, width=30, borderwidth=10)

    Label4 = Label(root, text='Insert the duration of insurance (annually)',font=font.Font(size=12),anchor='center')
    insurance_duration_entry = Entry(root, width=30, borderwidth=10)

    Label5 = Label(root, text='Insert the duration of deferred insurance (annually)',font=font.Font(size=12),anchor='center')
    insurance_deferred_entry = Entry(root, width=30, borderwidth=10)

    Label6 = Label(root, text='Insert the fractional annuity interval (annually)',font=font.Font(size=12),anchor='center')
    fractional_annuity_interval_entry = Entry(root, width=30, borderwidth=10)

    Label7 = Label(root, text='Insert the duration of annuity (annually)',font=font.Font(size=12),anchor='center')
    annuity_duration_entry = Entry(root, width=30, borderwidth=10)

    Label8 = Label(root,text='Insert the first benefit',font=font.Font(size=12),anchor='center')
    first_benefit_entry = Entry(root,width=30,borderwidth=10)

    Label9 = Label(root,text='Insert the second benefit',font=font.Font(size=12),anchor='center')
    second_benefit_entry = Entry(root,width=30,borderwidth=10)


    def insurance_validation():

        list_entry = [fractional_insurance_interval_entry, insurance_duration_entry, insurance_deferred_entry,
                      fractional_annuity_interval_entry,annuity_duration_entry,
                      first_benefit_entry,second_benefit_entry]
        
        if all(is_number(entry.get()) for entry in list_entry if not entry.cget("state") == "disabled"):

            global age,product,assumption,fractional_insurance
            global insurance_duration,insurance_deferred,fractional_annuity,annuity_duration
            global first_benefit,second_benefit

            product = product_str.get()
            age = float(((insurance_start_date-birth_date).days)/360)
            assumption = assumption_str.get()
            fractional_insurance= float(fractional_insurance_interval_entry.get())
            try : 
                insurance_duration = float(insurance_duration_entry.get())
            except ValueError :
                insurance_duration = 0
            try : 
                insurance_deferred = float(insurance_deferred_entry.get())
            except ValueError :
                insurance_deferred = 0

            fractional_annuity = float(fractional_annuity_interval_entry.get())
            annuity_duration = float(annuity_duration_entry.get())
            first_benefit = float(first_benefit_entry.get())

            try : 
                second_benefit = float(second_benefit_entry.get())
            except ValueError :
                second_benefit = 0 

            if (age +insurance_duration) > data.iloc[-1,0] or (age + fractional_insurance) > data.iloc[-1,0]:
                messagebox.showwarning('Invalid input','Insured age and insurance interval exceed max-age')
                fractional_insurance_interval_entry.delete(0,END)
                insurance_deferred_entry.delete(0,END)
                insurance_duration_entry.delete(0,END)

            elif (insurance_deferred > insurance_duration) :
                messagebox.showwarning('Invalid input','Insurance deferred can not be exceed insurance duration')
                fractional_insurance_interval_entry.delete(0,END)
                insurance_deferred_entry.delete(0,END)
                insurance_duration_entry.delete(0,END)

            elif (age +annuity_duration) > data.iloc[-1,0] or (age + fractional_annuity) > data.iloc[-1,0]:
                messagebox.showwarning('Invalid input','Insured age and annuity interval exceed max-age')
                fractional_annuity_interval_entry.delete(0,END)
                annuity_duration_entry.delete(0,END)

            elif (annuity_duration > insurance_duration) :
                messagebox.showwarning('Invalid input','Annuity duration can not be exceed insurance duration')
                insurance_duration_entry.delete(0,END)
                annuity_duration_entry.delete(0,END)

            elif (fractional_annuity > annuity_duration) or (fractional_insurance > insurance_duration ):
                messagebox.showwarning('Invalid input','Fractional interval should not exceed annuity/insurance duration')
                fractional_annuity_interval_entry.delete(0,END)
                fractional_insurance_interval_entry.delete(0,END)
                annuity_duration_entry.delete(0,END)
                insurance_duration_entry.delete(0,END)

            elif divmod(annuity_duration/fractional_annuity,1)[1] != 0:
                messagebox.showwarning('Invalid input','Annuities are not integer type')
                fractional_annuity_interval_entry.delete(0,END)
                annuity_duration_entry.delete(0,END)
            else :
                root.destroy()

        else :
            messagebox.showwarning('Invalid input','Input period must be in numeric format')
            fractional_insurance_interval_entry.delete(0,END)
            insurance_deferred_entry.delete(0,END)
            insurance_duration_entry.delete(0,END)
            fractional_annuity_interval_entry.delete(0,END)
            annuity_duration_entry.delete(0,END)
            first_benefit_entry.delete(0,END)
            second_benefit_entry.delete(0,END)

    button_get_data = Button(root,text='Submit',relief='raised',bg="#E0E0E0",command=insurance_validation)


    Label1.pack(fill=BOTH)
    insurance_product_option.pack(fill=BOTH)
    Label2.pack(fill=BOTH)
    assumption_option.pack(fill=BOTH)
    Label3.pack(fill=BOTH)
    fractional_insurance_interval_entry.pack(fill=BOTH)
    Label4.pack(fill=BOTH)
    insurance_duration_entry.pack(fill=BOTH)
    Label5.pack(fill=BOTH)
    insurance_deferred_entry.pack(fill=BOTH)
    Label6.pack(fill=BOTH)
    fractional_annuity_interval_entry.pack(fill=BOTH)
    Label7.pack(fill=BOTH)
    annuity_duration_entry.pack(fill=BOTH)
    Label8.pack(fill=BOTH)
    first_benefit_entry.pack(fill=BOTH)
    Label9.pack(fill=BOTH)
    second_benefit_entry.pack(fill=BOTH)
    button_get_data.pack(fill=BOTH)

    root.mainloop()
    return age,product,assumption,fractional_insurance,insurance_duration,insurance_deferred,fractional_annuity,annuity_duration,first_benefit,second_benefit



def create_entries(n,title):
    root = Tk()
    col_list = [0, 1, 2, 3, 4]
    row_list = np.arange(0, 2*(int(n/5)+1), 2)
    counter = 0
    entries = []
    list_value = []
    root.title(title)

    for i in row_list:
        for j in col_list:
            if counter < n:
                label = Label(root, text=f'Entry {counter + 1}')
                label.grid(row=i, column=j)
                entry = Entry(root)
                entry.grid(row=i+1, column=j)
                entries.append(entry)
                counter += 1

    def clicked():
        if checkbox_variable.get() == 1:
            list_true_entry = [bool(entry.get()) for entry in entries]
            last_index_true = None
            false_entry = []
            for index,value in enumerate(list_true_entry) :
                if value == True :
                    last_index_true = index
                if value == False :
                    false_entry.append(int(index))
            for i in false_entry :
                entries[i].insert(0,entries[last_index_true].get())
        else :
            pass

    def submit() :
        nonlocal list_value
        for entry in entries :
            try :
                list_value.append(float(entry.get()))
            except ValueError :
                messagebox.showwarning('Invalid input', 'Input period must be in numeric format')
                entry.delete(0,END)
                return
            
        if all(is_number(i) for i in list_value) == True:
            root.destroy()


    submit_button = Button(root,text='Submit',relief='raised',bg="#E0E0E0",command=submit)
    submit_button.grid(row=row_list[-1]+4, columnspan=5)
    checkbox_variable = IntVar()
    labelcheck = Label(root,text='Use the same interest rate for the rest')
    checkbox = Checkbutton(root, text="Yes", variable=checkbox_variable,command=clicked)
    labelcheck.grid(row=row_list[-1]+2,columnspan=2,sticky='w')
    checkbox.grid(row=row_list[-1]+3,columnspan=2,sticky='w')
    root.mainloop()

    return list_value



def last_option():
    def browse_folder():
        nonlocal folder_path
        folder_path = filedialog.askdirectory()
        if folder_path:
            root.destroy()

    folder_path = None

    root = Tk()
    root.title('Choose Your Folder Path')
    button = Button(root, relief='raised', bg="#E0E0E0", text="Choose Folder", command=browse_folder)
    button.pack(fill='both')

    root.mainloop()

    return folder_path

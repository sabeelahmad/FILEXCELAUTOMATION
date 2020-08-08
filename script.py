from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import pandas as pd
import json

root = Tk()
root.title('Excel Automation FIL')

# Get excel file location
root.filename = filedialog.askopenfilename(initialdir="/Desktop", title="Select Excel Sheet", filetypes=(("excel files", "*.xlsx"), ("all files", "*.*")))

# read selected file into a pandas dataframe
df = pd.read_excel(root.filename)
#grouped_df = df.groupby(by="Name of Supplier")

data = {}
counter = 1 

print(df.columns)

for index, row in df.iterrows():
  if not(row['Name of Supplier'] in data.keys()):
    data[row['Name of Supplier']] = {}
    # Add data for first time supplier comes across
    data[row['Name of Supplier']]['Address'] = row['Address']
    data[row['Name of Supplier']]['Type'] = row['Type']
    data[row['Name of Supplier']]['Rate/Kg'] = row['Rate/Kg']
    data[row['Name of Supplier']]['Nt Wt BQC'] = row['Nt Wt BQC']
    data[row['Name of Supplier']]['Nt Wt AQC'] = row['Nt Wt AQC']
    data[row['Name of Supplier']]['Unldg/WB/Tpt'] = row['Unldg/WB/Tpt']
  else:
    # Modify data if supplier already exists
    data[row['Name of Supplier']]['Nt Wt BQC'] += row['Nt Wt BQC']
    data[row['Name of Supplier']]['Nt Wt AQC'] += row['Nt Wt AQC']
    data[row['Name of Supplier']]['Unldg/WB/Tpt'] += row['Unldg/WB/Tpt']

# Calculate amount
for key, value in data.items():
  data[key]['Amount'] = data[key]['Nt Wt AQC']*data[key]['Rate/Kg']
  data[key]['Incentive'] = 0
  data[key]['Incentive Payment'] = 0
  data[key]['Payments'] = 0
  data[key]['Balance'] = 0

# Button handler
def set_data():
  supplier = clicked.get()
  incentive = float(e1.get())
  payment = float(e2.get())
  incentive_payment = float(e3.get())
  data[supplier]['Incentive'] = incentive
  data[supplier]['Incentive Payment'] = incentive_payment
  data[supplier]['Payments'] = payment 

  # clear fields
  e1.delete(0, END)
  e2.delete(0, END)
  e3.delete(0, END)
  
  # calculate balance payable
  data[supplier]['Balance'] = (data[supplier]['Amount']+data[supplier]['Incentive']) - (data[supplier]['Unldg/WB/Tpt'] + data[supplier]['Payments'] + data[supplier]['Incentive Payment'])
  
  # TODO : read data already in json file -> modify data for selected supplier 
  f = open('data.json')
  data2 = json.load(f)
  f.close()

  data2[supplier] = data[supplier]

  # update in json file
  with open('data.json', 'w') as fp:
    json.dump(data2, fp, indent=2)

# Dropdown of all supplier names
options = list(data.keys())
clicked = StringVar()
clicked.set(options[0]) # default value

dropdown = OptionMenu(root, clicked, *options)
dropdown.grid(row=0, column=0)

# Inputs for incentive and payment fields
e1 = Entry(root, width=30)
e1.insert(0, "Incentive: ")
e1.grid(row=1, column=0)

e3 = Entry(root, width=30)
e3.insert(0, "Incentive Payment: ")
e3.grid(row=1, column=1)

e2 = Entry(root, width=30)
e2.insert(0, "Payments: ")
e2.grid(row=1, column=2)

# Button to click and get values of fields
btn = Button(root, text="Add Data", command=set_data)
btn.grid(row=2, column=0)

# function that generates excel file
def generate_excel():
  gen_df = pd.DataFrame(columns=['Supplier', 'Address', 'Type', 'Nt Wt BQC', 'Nt Wt AQC', 'Rate/Kg', 'Amount',
  'Incentive', 'Incentive Payment', 'Unldg/WB/Tpt', 'Balance'])
  f = open('data.json')
  data = json.load(f)
  f.close()
  
  for k, v in data.items():
    new_row = {}
    new_row['Supplier'] = k
    for k2, v2 in v.items():
      new_row[k2] = v2
    gen_df = gen_df.append(new_row, ignore_index=True)
  
  # Save to excel file
  gen_df.to_excel(r'data.xlsx', sheet_name='Supplier Summary', index = False)
  messagebox.showinfo("Succesfull", "Excel File Generated Successfully!")

# Button to generate excel file
btn2 = Button(root, text="Generate Excel File", command=generate_excel)
btn2.grid(row=2, column=1)

root.mainloop()
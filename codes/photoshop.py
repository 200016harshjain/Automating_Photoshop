import win32com.client
import os
import pandas as pd

#Opening Photoshop and loading the PSD file
psApp = win32com.client.Dispatch("Photoshop.Application")
psApp.Open(r"C:\Users\ranja\Downloads\abhi.psd")
doc = psApp.Application.ActiveDocument


# creating a data frame and loading data from CSV

filename = r"C:\Users\ranja\Desktop\Scripts\student_data.csv"
df = pd.read_csv(filename)
print(df.head())

data_dict = df.to_dict('index')


#Manipulating the exact layers
layer_facts = doc.ArtLayers["NAME"]
text_of_layer = layer_facts.TextItem
text_of_layer.contents = "This is an example of a new text."

SIG_Name = doc.layerSets["TEXT"].ArtLayers["SIG_Name"]
sig_name_text = SIG_Name.TextItem
sig_name_text.contents = "random sig"

SMP_Name =  doc.layerSets["TEXT"].ArtLayers["SMP_Name"]
smp_name_text = SMP_Name.TextItem
smp_name_text.contents  = "random smp"


#Properties of like files we'll be making
options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
options.Format = 13   # PNG
options.PNG8 = False  # Sets it to PNG-24 bit

exportRoot = "C:/Users/ranja/Desktop/Scripts/"

# For every key and value in the items of the dictionary "dict_sample"
#loop the data and print
for dict_number,row in data_dict.items():
    
    # Replace Text of text Layer
    
    text_of_layer.contents = row['Name']
    smp_name_text.contents  = row['SMP']
    sig_name_text.contents = row['SIG']

    folder =  exportRoot + row['SIG'] + "/" +  row['SMP'] + "/" 
    if not os.path.exists(folder):
        os.makedirs(folder)
    fileName =  folder + row['Name'] + ".png"
    
    doc.Export(ExportIn=fileName, ExportAs=2, Options=options)

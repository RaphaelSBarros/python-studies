from tkinter import filedialog
import pandas as pd

file = filedialog.askopenfilename()

read_file = pd.read_excel (file) 

read_file.to_csv ("Test.csv",  index = None, header=True) 
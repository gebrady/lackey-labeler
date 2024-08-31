##### Suite Color Chooser Beta 1.0 #####
# Developed August-September 2022, Designed June 2022-Present
# Written by Graham E. Brady, with conceptual help from Rohan Shaiva and syntax help from Stack Overflow / Python docs
##### Defines Suite Color Chooser Class and Main functions when run from terminals #####
from tkinter import Tk, Label, Button, filedialog, Frame, PhotoImage
from PIL import ImageTk, Image
from tkinter.colorchooser import askcolor
import pandas as pd   
import os           
from pandastable import Table    
import time
from LabelMaker import LabelMaker
import random
import glob


###NEXT STEPS: ADD NEW LINE WITH EACH ADDED COLOR AND ITS SUITE<> instead of final line with summary
## Check for column names and standardize them
## check for format of lat and long to convert if necessary
## Add restart button and functions to stop the run with new filename
### Add a default when closing the file dialog and no selection is made

class ColorWheel(Tk):
  def __init__(self):
    super().__init__()
    p1 = PhotoImage(file='petrology.png')  
    self.iconphoto(False,p1)
    self.colors = []
    self.count = 0

    # configure the root window
    self.title('Suite Color Chooser')
    self.geometry('575x800')

    self.uploadButton = Button(self, text='Upload Data', command = self.UploadAction)
    self.uploadButton.pack()
    self.uploadLabel = Label(self, text="Click to upload data file.", fg="white")
    self.uploadLabel.pack()

    #self.logo = Label(self, image=p1)
    #self.logo.pack()
    
    """  # Buttons and Labels related to post data upload processes
    # button
    self.button = Button(self, text='Choose Color',command=self.callback)
    self.button.pack(pady=20)

    # label
    self.label = Label(self, text='Click the button to cycle through suites.\nThe first suite is: ' + self.suites[0], fg = "white")
    self.label.pack()
    
    # label 2
    self.label2 = Label(self, text='', fg = "white")
    self.label2.pack()
    """

  def __repr__(self):
    if self.colors:
      return str(self.colors)
    else: return 'Add some colors for the suites'

  def UploadAction(self):
    self.filename = filedialog.askopenfilename(initialdir=os.getcwd(), 
        title="Open Tabular Data File", 
        filetypes=[("Text files", "*.txt"),
        ("Comma-seperated values files", "*.csv"),
        ("Excel files", "*.xlsx")])   
    #self.label.configure(text = self.suites[self.count], fg = result[1])
    self.uploadLabel.configure(text= 'Selected: ' + str.split(self.filename,'/')[-1] + '. Check terminal for header.')
    if self.filename.endswith('.csv'): self.data = pd.read_csv(self.filename) # read data from dataTable
    elif self.filename.endswith('.xlsx'): self.data = pd.read_excel(self.filename)
    elif self.filename.endswith('.txt'): self.data = pd.read_table(self.filename)
    print(self.data)
    self.suites = self.data['Suite'].unique().tolist()
    #self.showTable(self.data)
    print('Shown above:', self.filename)
    #time.sleep(4)
    self.addColorTools()

  def addColorTools(self):
    # button
    self.button = Button(self, text='Choose Color',command=self.callback)
    self.button.pack(pady=20)

    # label
    self.label = Label(self, text='Click the "Choose Color" button to cycle through suites.\nThe first suite is: ' + self.suites[0], fg = "white")
    self.label.pack()
    
    # label 2
    self.label2 = Label(self, text='', fg = "white")
    self.label2.pack()

  def callback(self):
    result = askcolor(title = "colorwheel")
    self.label.configure(text = self.suites[self.count], fg = result[1])
    if self.count < len(self.suites)-1:
        self.label2.configure(text = "Choose next suite's color: " + self.suites[self.count+1])
    self.colors.append((self.suites[self.count], result[0]))
    self.count += 1
    print(result[1], " added to suite color list.")
    if self.count >= len(self.suites):
        self.label.configure(text = self.suites[-1], fg = result[1])
        self.label2.configure(text = "All suites' colors have been chosen!\n" + str(self.suites), fg = 'white')
        self.button2 = Button(self, text = 'Compile', command = self.compile).pack(pady=20)
  
  def compile(self):
      print(self.colors)
      a = LabelMaker('LabelTemplate.docx', self.data, self.colors)
      a.writeTable(os.path.join(os.getcwd(),str.split(str.split(self.filename,'/')[-1],'.')[0])) # add output filename here based on original file minus leading dirs and extension
      print('File written to: ' + os.path.join(os.getcwd(),str.split(str.split(self.filename,'/')[-1],'.')[0]) + '_Labels.docx')
      self.frame = Frame(self, width=300, height=300)
      self.frame.pack(after=self.button2)
      #self.frame.place(anchor='center', relx=0.5, rely=0.5)

      # Create an object of tkinter ImageTk
      imgfile = random.choice(glob.glob('./funimg/*.jpg')) #choose random photo from the folder
      self.img = ImageTk.PhotoImage(Image.open(imgfile))

      # Create a Label Widget to display the text or Image
      self.isodope = Label(self.frame, image = self.img)
      self.isodope.pack(pady=10)

      self.writtenLabel = Label(self, text= 'Your labels were written to ' + os.path.join(os.getcwd(),str.split(str.split(self.filename,'/')[-1],'.')[0]) + '_Labels.docx')
      self.writtenLabel.pack()

      self.button3 = Button(self, text = 'Quit', command = self.destroy).pack(pady=20)


if __name__ == "__main__":
  app = ColorWheel()
  app.mainloop()

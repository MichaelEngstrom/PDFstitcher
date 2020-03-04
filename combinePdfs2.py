""" combinePdfs - Michael Engstrom

    GUI to combine separate PDFs into single PDF
    File dialog to select target PDFs
    PDF paths saved one per line in list
    Read each path in list and combine into one PDF
    Count of PDFs and remaining and number combined

    Developed using Python 3.6

    Install PyPDF2 module in python environment using pip:

    pip install PyPDF2

    This allows import of PyPDF2 module in script """

import tkinter as tk
import os, PyPDF2
from tkinter import filedialog
from tkinter import StringVar
from tkinter import messagebox

# Create list to hold script paths
pdfs_list = []

class PdfCombiner:
    def __init__(self):
        # Create the main window
        self.main_window = tk.Tk()
        self.main_window.title("Combine Multiple PDFs into Single PDF")

        # Create the frames
        self.button_frame = tk.Frame(self.main_window)
        self.pdflabel_frame = tk.Frame(self.main_window)
        self.textwindow_frame = tk.Frame(self.main_window)
        self.numpdfs_frame = tk.Frame(self.main_window)
        self.combpdfs_frame = tk.Frame(self.main_window)

        # Create and pack widgets for button frame
        self.addpdf_button = tk.Button(self.button_frame, \
                                        text='Add PDF(s)', \
                                        command=self.add_pdf)
        self.removepdf_button = tk.Button(self.button_frame, \
                                        text='Remove Last PDF',\
                                        command=self.remove_pdf)
        self.removeallpdfs_button = tk.Button(self.button_frame, \
                                        text='Remove All PDFs',\
                                        command=self.removeall_pdfs)
        self.combine_button = tk.Button(self.button_frame, \
                                        text='Combine PDFs Now!',\
                                        command=self.combine_pdfs)
        self.quit_button = tk.Button(self.button_frame, \
                                     text='Quit', \
                                     command=self.main_window.destroy)
        self.addpdf_button.pack(side=tk.LEFT, pady=10, padx=20)
        self.removepdf_button.pack(side=tk.LEFT, pady=10, padx=20)
        self.removeallpdfs_button.pack(side=tk.LEFT, pady=10, padx=20)
        self.combine_button.pack(side=tk.LEFT, pady=10, padx=20)
        self.quit_button.pack(side=tk.LEFT, pady=10, padx=20)

        # Create and pack widget for test label frame
        self.tests_label = tk.Label(self.pdflabel_frame, \
                                  text='List of PDFs to be combined:')
        self.tests_label.pack(side=tk.LEFT)

        # Create and pack widget for text window frame
        self.pdfs_text = tk.Text(self.textwindow_frame, \
                                  height=15, width=100)
        self.pdfs_text.pack(expand='YES', fill=tk.BOTH, pady=3)

        # Create and pack widgets for number of tests frame
        self.pdfstocombine_label = tk.Label(self.numpdfs_frame, \
                                         text='PDFs Selected: ')
        self.runcount = tk.StringVar()  # To update numpdfs_label
        self.numpdfs_label = tk.Label(self.numpdfs_frame, \
                                       textvariable=self.runcount)
        self.pdfstocombine_label.pack(side=tk.LEFT)
        self.numpdfs_label.pack(side=tk.LEFT)

        # Create and pack widgets for tests completed frame
        self.testcomp_label = tk.Label(self.combpdfs_frame, \
                                       text='PDFs Combined: ')
        self.completed = tk.StringVar() # To update numcomp_label
        self.numcomp_label = tk.Label(self.combpdfs_frame, \
                                      textvariable=self.completed)
        self.testcomp_label.pack(side=tk.LEFT)
        self.numcomp_label.pack(side=tk.LEFT)

        # Pack the frames
        self.button_frame.pack()
        self.pdflabel_frame.pack()
        self.textwindow_frame.pack()
        self.numpdfs_frame.pack()
        self.combpdfs_frame.pack()

        # Start the main loop
        tk.mainloop()

    # The add_pdf method is the callback function for the addpdf_button widget
    def add_pdf(self):
        # Use file dialog to add file path to list one at a time
        fd = tk.filedialog.askopenfilenames(filetypes =(("PDF Files","*.pdf*"),
                                                       ("All Files","*.*")),
                                           title = "Choose PDF File(s)")
        if fd:
            # Copy path to string then append to path list
            for filepath in fd:
                pdfs_list.append(filepath)

            # Write changes to text box
            self.update_text_box()


    # The remove_pdf method pops the last pdf in list, writes to text box
    def remove_pdf(self):
        if pdfs_list:    #if not empty
            pdfs_list.pop()
            self.update_text_box()
    
    # The remove_pdf method pops the last pdf in list, writes to text box
    def removeall_pdfs(self):
        pdfs_list.clear()

        # Clear box 
        self.clear_text_box()

        # Remove queued and combined counts
        self.runcount.set("")
        self.completed.set("")

    # The combine_pdfs method is the callback function for combine_button widget
    def combine_pdfs(self):
        
        # File dialog for output PDF of combined PDFs
        if len(pdfs_list) > 0:
            fd = tk.filedialog.asksaveasfilename(initialdir = "/",
                                               title = "Specify PDF Output",
                                               filetypes = (("PDF files","*.pdf"),
                                                            ("all files","*.*")))

            # Clear the completed count
            completed_count = 0
            
            # Combine the PDFs
            for pdf in pdfs_list:
                    
                # Call the script

                pdfs_list.sort(key = str.lower)

                pdfWriter = PyPDF2.PdfFileWriter()
                
                # Loop through all the PDF files.
                for filepath in pdfs_list:
                    pdfFileObj = open(filepath, 'rb')
                    pdfReader = PyPDF2.PdfFileReader(pdfFileObj)
                    
                    # Loop through all the pages (except the first) and add them.
                    for pageNum in range(1, pdfReader.numPages):
                        pageObj = pdfReader.getPage(pageNum)
                        pdfWriter.addPage(pageObj)

                # Save the resulting PDF to a file.
                pdfOutput = open(fd, 'wb')
                pdfWriter.write(pdfOutput)
                pdfOutput.close()

                # Update Combined PDFs count
                completed_count += 1
                self.completed.set(str(completed_count))

            # Message box alerting user operation succeeded
            pdfs_combined = str(completed_count) + " PDFs combined as: " + fd
            messagebox.showinfo("Operation Complete", pdfs_combined)

        else:
            messagebox.showinfo("Information", "No PDF(s) selected")

    # Method to update text box
    def update_text_box(self):
        # Clear text box for writing appended list
        self.clear_text_box()
        
        # Write list of pdfs to text box
        for path in pdfs_list:
            # Append path to line in text box
            self.pdfs_text.insert(tk.END, path)
            self.pdfs_text.insert(tk.END, '\n')
            
        # Update 'PDFs to combine' count in GUI for number of pdfs selected
        self.runcount.set(len(pdfs_list))
        self.completed.set("")

    def clear_text_box(self):
        # Clear text box for writing appended list
        self.pdfs_text.delete('1.0', tk.END)
    

# Create an instance of the PdfCombiner class
pdf_combine = PdfCombiner()
    

# General imports
import Tkinter as Tk
import ttk
import tkFileDialog
import tkMessageBox
import os
import logging
import math
import webbrowser
from PIL import Image, ImageTk

# Package specific imports
import load_est
import load_est.constants as constants


class MainGUI():

	def __init__(self):

		self.abort = False

		title = 'PSC Load Esitmate Tool'
		# Initialise constants and Tk window
		self.master = Tk.Tk()
		self.master.title(title)

		# Ensure that on_closing is processed correctly
		self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

		self.master.mainloop()


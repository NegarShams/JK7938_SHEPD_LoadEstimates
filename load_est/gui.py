"""
#######################################################################################################################
###											PSSE G74 Fault Studies													###
###		Script sets up PSSE to carry out fault studies in line with requirements of ENA G74							###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""
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
from collections import OrderedDict

# Package specific imports
import load_est
import load_est.constants as constants
import Load_Estimates_to_PSSE


class MainGUI:
	"""
		Main class to produce the GUI
		Allows the user to select the busbars and methodology to be applied in the fault current calculations
	"""
	def __init__(self, title=constants.GUI.gui_name, station_dict=dict()):
		"""
			Initialise GUI
		:param str title: (optional) - Title to be used for main window
		:param str sav_case: (optional) - Full path to the existing SAV case
		:param list busbars:  (optional) - List of busbars selected from slider
		TODO: (low priority) Potential option for parallel processing where PSSE prepares the model once SAV case selected
		TODO: 	and user is still selecting busbars but very hard to implement.
		TODO: (low priority) Additional parameters options to select some constants (kA vs A output)
		"""
		# Get logger handle
		self.logger = logging.getLogger(constants.Logging.logger_name)
		self.abort = False

		# Initialise constants and Tk window
		self.master = Tk.Tk()
		self.master.title(title)
		# Ensure that on_closing is processed correctly
		self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

		self.fault_times = list()
		# General constants which need to be initialised
		self._row = 0
		self._col = 0

		# Target file that results will be exported to
		self.target_file = str()
		self.results_pth = os.path.dirname(os.path.realpath(__file__))

		# Stand alone command button constants
		self.cmd_select_sav_case = Tk.Button()

		# PSC logo constants
		self.hyp_help_instructions = Tk.Label()
		self.psc_logo_wm = Tk.PhotoImage()
		self.psc_logo = Tk.Label()
		self.psc_info = Tk.Label()

		# Load Options label frame constants
		self.load_labelframe = ttk.LabelFrame()
		self.load_radio_opt_sel = Tk.IntVar()
		self.load_radio_btn_list = list()
		self.load_radio_opts = OrderedDict()
		self.load_prev_radio_opt = int()

		# year drop down options
		self.year_selected = Tk.StringVar()
		if bool(station_dict):
			constants.GUI.year_list = sorted(station_dict[0].load_forecast_dict.keys())
			constants.GUI.season_list = sorted(station_dict[0].seasonal_percent_dict.keys())
			# constants.GUI.station_dict = station_dict
		self.year_list = constants.GUI.year_list
		self.year_om = None

		# season drop down options
		self.season_selected = Tk.StringVar()
		self.season_list = constants.GUI.season_list
		self.season_om = None

		# Load Selector scroll constants
		self.load_side_lbl = Tk.Label()
		self.load_select_frame = None
		self.load_canvas = Tk.Canvas()
		self.load_entry_frame = ttk.Frame()
		self.load_scrollbar = ttk.Scrollbar()
		# Load selector zone constants
		self.load_boolvar_zon = dict()
		self.zones = dict()
		# todo just for testing to set a zone dict, set equal to constant
		for i in range(13):
			self.zones[i] = "zone " + str(i)
		# Load selector gsp constants
		self.load_boolvar_gsp = dict()
		self.load_gsps = dict()
		# todo just for testing to set a gsp dict, set equal to constant
		for i in range(28):
			self.load_gsps[i] = "gsp " + str(i)

		# Add PSC logo with hyperlink to the website
		self.add_psc_logo(row=self.row(), col=0)

		self.cmd_import_load_estimates = self.add_cmd_button(
			label='Import SSE Load Estimates .xlsx', cmd=self.import_load_estimates, row=self.row(), col=3)

		# Add PSC logo with hyperlink to the website
		self.add_psc_logo(row=self.row(), col=5)

		# SAV case to be run on
		self.sav_case = str()

		# Determine label for button based on the SAV case being loaded
		if self.sav_case:
			lbl_sav_button = 'SAV case = {}'.format(os.path.basename(self.sav_case))
		else:
			lbl_sav_button = 'Select PSSE SAV Case'

		self.cmd_select_sav_case = self.add_cmd_button(
			label=lbl_sav_button, cmd=self.select_sav_case,  row=self.row(1), col=3)

		# create load options labelframe
		self.create_load_options()

		#
		self._col, self._row, = self.master.grid_size()
		# add cmd button to scale load
		self.cmd_scale_load_gen = self.add_cmd_button(
			label='Scale Load and Generation',
			cmd=self.scale_loads,
			row=self.row(1), col=3
		)
		self.cmd_scale_load_gen.configure(state=Tk.DISABLED)

		# # Add tick box for whether it needs to be opened again on completion
		# self.add_open_excel(row=self.row(1), col=self.col())

		# Add PSC logo in Windows Manager
		self.add_psc_logo_wm()

		# Add PSC UK and phone number
		self.add_psc_phone(row=self.row(1), col=0)

		# Generation label frame constants
		self.gen_labelframe = ttk.LabelFrame()
		self.gen_radio_opt_sel = Tk.IntVar()
		self.gen_radio_btn_list = list()
		self.gen_radio_opts = OrderedDict()
		self.gen_prev_radio_opt = int()

		# Generation Selector scroll constants
		self.gen_side_lbl = Tk.Label()
		self.gen_select_frame = None
		self.gen_canvas = Tk.Canvas()
		self.gen_entry_frame = ttk.Frame()
		self.gen_scrollbar = ttk.Scrollbar()
		# gen selector zone constants
		self.gen_boolvar_zon = dict()

		# create load options labelframe
		self.create_gen_options()

		self.logger.debug('GUI window created')
		# Produce GUI window
		self.master.mainloop()

	def create_load_options(self):

		self.load_labelframe = ttk.LabelFrame(self.master, text='Load Scaling Options')
		self.load_labelframe.grid(row=4, column=0, columnspan=4, sticky='NSEW')

		self.load_radio_opt_sel = Tk.IntVar(self.master, 0)
		self.load_prev_radio_opt = 1

		# Dictionary to create multiple buttons
		self.load_radio_opts = OrderedDict([
			('None', 0),
			('All Loads', 1),
			('Selected GSP', 2),
			('Selected Zones', 3)
		])

		self._row = 0
		self.load_radio_btn_list = list()
		# Loop is used to create multiple Radiobuttons
		# rather than creating each button separately
		for (text, value) in self.load_radio_opts.items():
			temp_rb = Tk.Radiobutton(
				self.load_labelframe,
				text=text,
				variable=self.load_radio_opt_sel,
				value=value,
				command=self.load_radio_button_click
			)
			temp_rb.grid(row=self.row(), sticky='W', padx=20, pady=5)
			self.load_radio_btn_list.append(temp_rb)
			self.row(1)

		self.enable_radio_buttons(self.load_radio_btn_list, enable=False)

		# add drop down list for year
		self.year_selected = Tk.StringVar(self.master)
		self.year_list = constants.GUI.year_list
		self.year_selected.set(self.year_list[0])

		self.year_om = self.add_drop_down(
			row=self.row(),
			col=self.col(),
			var=self.year_selected,
			list=self.year_list,
			location=self.load_labelframe
		)
		self.year_om.configure(state=Tk.DISABLED)
		#
		# add drop down list for season
		self.season_selected = Tk.StringVar(self.master)
		self.season_list = constants.GUI.season_list
		self.season_selected.set(self.season_list[0])

		self.season_om = self.add_drop_down(
			row=self.row(1), col=self.col(),
			var=self.season_selected,
			list=self.season_list,
			location=self.load_labelframe
		)
		self.season_om.configure(state=Tk.DISABLED)

		return None

	def create_gen_options(self):

		self.gen_labelframe = ttk.LabelFrame(self.master, text='Generation Scaling Options')
		self.gen_labelframe.grid(row=4, column=4, columnspan=4, sticky='NSEW')

		self.gen_radio_opt_sel = Tk.IntVar(self.master, 0)
		self.gen_prev_radio_opt = 1

		# Dictionary to create multiple buttons
		self.gen_radio_opts = OrderedDict([
			('None', 0),
			('All Generators', 1),
			('Selected Zones', 2)
		])

		self._row = 0
		self.gen_radio_btn_list = list()
		# Loop is used to create multiple Radiobuttons
		# rather than creating each button separately
		for (text, value) in self.gen_radio_opts.items():
			temp_rb = Tk.Radiobutton(
				self.gen_labelframe,
				text=text,
				variable=self.gen_radio_opt_sel,
				value=value,
				command=self.gen_radio_button_click
			)
			temp_rb.grid(row=self.row(), sticky='W', padx=20, pady=5)
			self.gen_radio_btn_list.append(temp_rb)
			self.row(1)

		self.enable_radio_buttons(self.gen_radio_btn_list, enable=False)

		return None

	def row(self, i=0):
		"""
			Returns the current row number + i
		:param int i: (optional=0) - Will return the current row number + this value
		:return int _row:
		"""
		self._row += i
		return self._row

	def col(self, i=0):
		"""
			Returns the current col number + i
		:param int i: (optional=0) - Will return the current col number + this value
		:return int _row:
		"""
		self._col += i
		return self._col

	def add_drop_down(self, row, col, var, list, location=None):
		"""
			Function to all a list of optimisation options which will become enabled if the user selects to run a
			virtual statcom study
		:param int row: Row number to use
		:param int col: Column number to use
		:rtype Tk.OptionMenu
		:return dropdown_optimisaiton_option:
		"""
		# Check whether there is a successfully loaded SAV case to enable the list option
		if location is None:
			var.set(list[0])
			# Create the drop down list to be shown in the GUI
			w = Tk.OptionMenu(
				self.master,
				var,
				*list
			)
			w.grid(row=row, column=col, columnspan=2, sticky=Tk.W + Tk.E)
		else:
			var.set(list[0])
			# Create the drop down list to be shown in the GUI
			w = Tk.OptionMenu(
				location,
				var,
				*list
			)
			w.grid(row=row, column=col, columnspan=2, sticky=Tk.W + Tk.E, padx=20, pady=5)

		# return dropdown_optimisation_options
		return w

	def add_cmd_button(self, label, cmd, row, col, location=None):
		"""
			Function just adds the command button to the GUI which is used for selecting the SAV case
		:param int row: Row number to use
		:param int col: Column number to use
		:return None:
		"""
		if location is None:
			# Create button and assign to Grid
			cmd_btn = Tk.Button(self.master, text=label, command=cmd)
			# self.cmd_select_sav_case.grid(row=row, column=col, columnspan=2, sticky=Tk.W + Tk.E)
			cmd_btn.grid(row=row, column=col, columnspan=2, sticky=Tk.W + Tk.E, padx=5, pady=5)
			# CreateToolTip(widget=self.cmd_select_sav_case, text=(
			# 	'Select the SAV case for which fault studies should be run.'
			# ))
		else:
			cmd_btn = Tk.Button(location, text=label, command=cmd)
			# self.cmd_select_sav_case.grid(row=row, column=col, columnspan=2, sticky=Tk.W + Tk.E)
			cmd_btn.grid(row=row, column=col, columnspan=2, sticky=Tk.W + Tk.E, padx=5, pady=5)
			# CreateToolTip(widget=self.cmd_select_sav_case, text=(
			# 	'Select the SAV case for which fault studies should be run.'
			# ))

		return cmd_btn

	def scale_loads(self):

		return

	def import_load_estimates(self):

		return

	def load_radio_button_click(self):

		if self.load_radio_opt_sel.get() == 0:
			self.year_om.config(state=Tk.DISABLED)
			self.season_om.config(state=Tk.DISABLED)
			self.remove_load_bottom_selector()
		if self.load_radio_opt_sel.get() == 1:
			self.year_om.config(state=Tk.NORMAL)
			self.season_om.config(state=Tk.NORMAL)
			self.remove_load_bottom_selector()
		elif self.load_radio_opt_sel.get() > 1 and self.load_radio_opt_sel.get() != self.load_prev_radio_opt:
			self.year_om.config(state=Tk.NORMAL)
			self.season_om.config(state=Tk.NORMAL)
			self.create_load_select_frame()

		self.load_prev_radio_opt = self.load_radio_opt_sel.get()
		return None

	def gen_radio_button_click(self):

		if self.load_radio_opt_sel.get() == 0:
			self.year_om.config(state=Tk.DISABLED)
			self.season_om.config(state=Tk.DISABLED)
			self.remove_load_bottom_selector()
		if self.load_radio_opt_sel.get() == 1:
			self.year_om.config(state=Tk.NORMAL)
			self.season_om.config(state=Tk.NORMAL)
			self.remove_load_bottom_selector()
		elif self.load_radio_opt_sel.get() > 1 and self.load_radio_opt_sel.get() != self.load_prev_radio_opt:
			self.year_om.config(state=Tk.NORMAL)
			self.season_om.config(state=Tk.NORMAL)
			self.create_load_select_frame()

		self.load_prev_radio_opt = self.load_radio_opt_sel.get()
		return None

	def enable_radio_buttons(self, radio_btn_list, enable=True):

		for radio_btn in radio_btn_list:
			if enable:
				radio_btn.config(state=Tk.NORMAL)
			else:
				radio_btn.config(state=Tk.DISABLED)

		return None

	def create_load_select_frame(self):

		if self.load_radio_opt_sel.get() == 2:
			lbl = 'Select GSP(s):'
		elif self.load_radio_opt_sel.get() == 3:
			lbl = 'Select Zone(s):'

		if self.load_select_frame is None:
			master_col, master_rows,  = self.master.grid_size()
			self._row = master_rows
			self._col = 0
			# Label for what is included in entry
			self.load_side_lbl = Tk.Label(master=self.load_labelframe, text=lbl)
			self.load_side_lbl.grid(row=self.row(1), column=self.col())

			# Produce a frame which will house the zones check buttons and place them in the window
			# self.zon_frame = ttk.Frame(self.master, relief=Tk.GROOVE, style=self.styles.frame_outer)
			self.load_select_frame = ttk.Frame(self.load_labelframe, relief=Tk.GROOVE)
			# zon_frame = ttk.Frame(self.master, relief=Tk.GROOVE, bd=1, style=self.styles.frame)
			self.load_select_frame.grid(row=self.row(1), column=self.col(), columnspan=4)

			# Create a canvas into which the zon_frame will be housed so that scroll bars can be added for the
			# network zone list
			self.load_canvas = Tk.Canvas(self.load_select_frame)

			# Add the canvas into a new frame
			# self.entry_frame = ttk.Frame(self.canvas, style=self.styles.frame)
			self.load_entry_frame = ttk.Frame(self.load_canvas)

			# Create scroll bars which will control the zon_frame within the canvas and configure the controls
			# self.zon_scrollbar = ttk.Scrollbar(
			# 	self.zon_frame, orient="vertical", command=self.canvas.yview, style=self.styles.scrollbar
			# )
			self.load_scrollbar = ttk.Scrollbar(
				self.load_select_frame, orient="vertical", command=self.load_canvas.yview
			)
			self.load_canvas.configure(yscrollcommand=self.load_scrollbar.set)

			# Locate the scroll bars on the right hand side of the canvas and locate canvas within newly created frame
			self.load_scrollbar.pack(side="right", fill="y")
			self.load_canvas.pack(side="left")
			self.load_canvas.create_window((0, 0), window=self.load_entry_frame, anchor='nw')

			# Bind the action of the scrollbar with the movement of the canvas
			self.load_entry_frame.bind("<Configure>", self.canvas_scroll_function)
			self.move_widgets_down()

		else:
			self.load_side_lbl.config(text=lbl)
			if len(self.load_entry_frame.winfo_children()) > 0:
				for widget in self.load_entry_frame.winfo_children():
					widget.destroy()

		if self.load_radio_opt_sel.get() == 2:
			# 'Select GSP(s):'
			for gsp in self.load_gsps:
				self.load_boolvar_gsp[gsp] = Tk.BooleanVar()
				self.load_boolvar_gsp[gsp].set(0)

			counter = 0
			for gsp in self.load_gsps:
				lbl = "{}".format(self.load_gsps[gsp])
				check_button = ttk.Checkbutton(
					self.load_entry_frame, text=lbl, variable=self.load_boolvar_gsp[gsp],
				)
				check_button.grid(row=counter, column=0, sticky="w")
				counter += 1
			self.load_canvas.yview_moveto(0)

		if self.load_radio_opt_sel.get() == 3:
			# 'Select Zone(s):'
			for zone in self.zones:
				self.load_boolvar_zon[zone] = Tk.BooleanVar()
				self.load_boolvar_zon[zone].set(0)

			counter = 0
			for zone in self.zones:
				lbl = "{}".format(self.zones[zone])
				check_button = ttk.Checkbutton(
					self.load_entry_frame, text=lbl, variable=self.load_boolvar_zon[zone],
				)
				check_button.grid(row=counter, column=0, sticky="w")
				counter += 1
			self.load_canvas.yview_moveto(0)

		return None

	def canvas_scroll_function(self, _event):
		"""
			Function to control what happens when the frame is scrolled
		:return None:
		"""

		# self.canvas.configure(
		# 	scrollregion=self.canvas.bbox("all"), width=230, height=200, background=self.styles.bg_color_frame)
		self.load_canvas.configure(
			scrollregion=self.load_canvas.bbox("all"), width=230, height=200)
		return None

	def remove_load_bottom_selector(self):

		if self.load_select_frame is not None:
			self.load_side_lbl.grid_remove()
			self.load_select_frame.grid_remove()
			self.load_side_lbl = None
			self.load_select_frame = None

	def move_widgets_down(self):

		master_col, master_rows, = self.master.grid_size()
		self.cmd_scale_load_gen.grid(row=master_rows + 1, column=3)
		self.psc_info.grid(row=master_rows+2, column=0)

		return None

	def add_entry_fault_times(self, row, col):
		"""
			Function to add the text entry row for inserting fault times
		:param int row:  Row number to use
		:param int col:  Column number to use
		:return None:
		"""
		# Label for what is included in entry
		lbl = Tk.Label(master=self.master, text='Desired Fault Times\n(in seconds separated by commas)')
		lbl.grid(row=row, column=col, rowspan=2, sticky=Tk.W + Tk.N + Tk.S)
		# Set initial value for variable
		self.var_fault_times_list.set(constants.GUI.default_fault_times)
		# Add entry box
		self.entry_fault_times = Tk.Entry(master=self.master, textvariable=self.var_fault_times_list)
		self.entry_fault_times.grid(row=row, column=col + 1, sticky=Tk.W + Tk.E, rowspan=2)
		CreateToolTip(widget=self.entry_fault_times, text=(
			'Enter the durations after the fault the current should be calculated for.\n'
			'Multiple values can be input in a list.'
		))
		return None

	def add_reload_sav(self, row, col):
		"""
			Function to add a tick box on whether the user wants to reload the SAV case
		:param int row:  Row number to use
		:param int col:  Column number to use
		:return None:
		"""
		lbl = 'Reload initial SAV case on completion'
		self.bo_reload_sav.set(constants.GUI.reload_sav_case)
		# Add tick box
		check_button = Tk.Checkbutton(
			self.master, text=lbl, variable=self.bo_reload_sav
		)
		check_button.grid(row=row, column=col, columnspan=2, sticky=Tk.W)
		CreateToolTip(widget=check_button, text=(
			'If selected the SAV case will be reloaded at the end of this study, if not then the model will as the '
			'study finished which may be useful for debugging purposes.'
		))
		return None

	def add_fault_types(self, col):
		"""
			Function to add a tick box to select the available fault types
		:param int col:  Column number to use
		:return None:
		"""
		labels = (
			'3 Phase fault (BKDY method)',
			'3 Phase fault (IEC method)',
			'LG Phase fault (IEC method)'
		)
		boolean_vars = (self.bo_fault_3_ph_bkdy, self.bo_fault_3_ph_iec, self.bo_fault_1_ph_iec)

		# TODO: Implement additional fault current calculations
		# The following values are used to enable and disable the buttons and will be removed once
		# the fault current implementations for IEC has been developed.
		default_values = (1, 0, 0)
		enabled = (True, False, False)
		i = 0
		for i, lbl in enumerate(labels):
			# Defaults assuming that all faults will be calculated
			boolean_vars[i].set(default_values[i])
			# Add check button for this fault
			check_button = Tk.Checkbutton(
				self.master, text=lbl, variable=boolean_vars[i]
			)
			check_button.grid(row=self.row(1), column=col, sticky=Tk.W)
			# TODO: Temporary to disable non-necessary faults
			# Disable faults that are not important
			if not enabled[i]:
				check_button.config(state='disabled')
		return i

	def add_open_excel(self, row, col):
		"""
			Function to add a tick box on whether the user wants to open the Excel file of results at the end
			:param int row:  Row number to use
			:param int col:  Column number to use
			:return None:
		"""
		lbl = 'Open exported Excel file'
		self.bo_open_excel.set(constants.GUI.open_excel)
		# Add tick box
		check_button = Tk.Checkbutton(
			self.master, text=lbl, variable=self.bo_open_excel
		)
		check_button.grid(row=row, column=col, columnspan=2, sticky=Tk.W)
		CreateToolTip(widget=check_button, text=(
			'If selected the exported excel file will be loaded and visible on completion of the study.'
		))
		return None

	def add_hyp_help_instructions(self, row, col):
		"""
			Function just adds the hyperlink to the GUI which is used for loading the work instructions
		:param int row: Row number to use
		:param int col: Column number to use
		:return: None
		"""
		# Create Help link and reference to the work instructions document
		self.hyp_help_instructions = Tk.Label(self.master, text='Help Instructions', fg='Blue', cursor='hand2')
		self.hyp_help_instructions.grid(row=row, column=col, sticky=Tk.W)
		self.hyp_help_instructions.bind('<Button - 1>', lambda e: webbrowser.open_new(constants.GUI.local_directory + '\\JK7938-01-00 PSSE G74 Fault Current Tool - Work Instruction.pdf'))
		return None

	def add_psc_logo_wm(self):
		"""
			Function just adds the PSC logo to the windows manager in GUI
		:return: None
		"""
		# Create the PSC logo for including in the windows manager
		self.psc_logo_wm = Tk.PhotoImage(file=constants.GUI.img_pth_window)
		self.master.tk.call('wm', 'iconphoto', self.master._w, self.psc_logo_wm)
		return None

	def add_psc_info(self, row, col):
		"""
			Function just adds the PSC company info and contact details
		:param row: Row number to use
		:param col: Column number to use
		:return: None
		"""
		# Create the PSC company info and contact details
		self.psc_info = Tk.Label(
			self.master, text=constants.GUI.psc_uk, justify='center', font=constants.GUI.psc_font,
			foreground=constants.GUI.psc_color_web_blue
		)
		self.psc_info.grid(row=row, column=col, columnspan=2)
		return None

	def add_psc_logo(self, row, col):
		"""
			Function just adds the PSC logo in the GUI and a hyperlink to the website
		:param row: Row number to use
		:param col: Column Number to use
		:return: None
		"""
		# Create the PSC logo and a hyperlink to the website
		# #img = Image.open(constants.GUI.local_directory + '\\PSC_logo.gif').resize((35,35), Image.ANTIALIAS)
		# #img = Image.open(constants.GUI.img_pth).resize((35, 35), Image.ANTIALIAS)
		img = Image.open(constants.GUI.img_pth_main)
		img.thumbnail(constants.GUI.img_size)
		img = ImageTk.PhotoImage(img)
		# #self.psc_logo = Tk.Label(self.master, image=img, text = 'www.pscconsulting.com', cursor = 'hand2', justify = 'center', compound = 'top', fg = 'blue', font = 'Helvetica 7 italic')
		self.psc_logo = Tk.Label(self.master, image=img, cursor='hand2', justify='center', compound='top')
		self.psc_logo.photo = img
		self.psc_logo.grid(row=row, column=col, columnspan=3, rowspan=3)
		self.psc_logo.bind('<Button - 1>', lambda e: webbrowser.open_new('https://www.pscconsulting.com/'))
		return None

	def add_psc_phone(self, row, col):
		"""
			Function just adds the PSC contact details
		:param row: Row number to use
		:param col: Column number to use
		:return: None
		"""
		# Create the PSC company info and contact details
		self.psc_info = Tk.Label(
			self.master, text=constants.GUI.psc_phone, justify='center', font=constants.GUI.psc_font,
			foreground=constants.GUI.psc_color_grey
		)
		self.psc_info.grid(row=row, column=col, columnspan=2)
		return None

	def add_sep(self, row, col_span):
		"""
			Function just adds a horizontal separator
		:param int row: Row number to use
		:param int col_span: Column span number to use
		:return None:
		"""
		# Add separator
		sep = ttk.Separator(self.master, orient="horizontal")
		sep.grid(row=row, sticky=Tk.W + Tk.E, columnspan=col_span, pady=5)
		return None

	def import_busbars_list(self):
		"""
			Function to import a list of busbars based on the selected file
		:return: None
		"""
		# Ask user to select file(s) or folders based on <.bo_files>
		file_path = tkFileDialog.askopenfilename(
			initialdir=self.results_pth,
			filetypes=constants.General.file_types,
			title='Select spreadsheet containing list of busbars'
		)

		# Import busbar list from file assuming it is first column and append to existing list
		busbars = load_est.file_handling.import_busbars_list(path=file_path)
		self.selected_busbars.extend(busbars)

		# Update results path to include this name
		self.results_pth = os.path.dirname(file_path)

		return None

	def edit_busbars_list(self):
		"""
			Function to popup a list of busbars that are to be faulted so that new / different busbars can be
			added or removed from the list
		:return None:
		"""
		# TODO: Write a pop-up window that will allow busbars list to be edited / manually populated
		# Create new window to house the busbars list and ensure it pops up on top of everything else
		busbars_window = Tk.Toplevel(self.master)
		busbars_window.attributes('-topmost', 'true')
		bus_data = BusbarsWindow(master=busbars_window, busbars=self.selected_busbars)
		# Add entry boxes and populate busbars
		bus_data.add_entry_boxes()
		bus_data.populate_busbars()

	def select_sav_case(self):
		"""
			Function to allow the user to select the SAV case to run
		:return: None
		"""
		# Ask user to select file(s) or folders based on <.bo_files>
		file_path = tkFileDialog.askopenfilename(
			initialdir=self.results_pth,
			filetypes=constants.General.sav_types,
			title='Select SAV case for fault studies'
		)

		# Import busbar list from file assuming it is first column and append to existing list
		self.sav_case = file_path
		lbl_sav_button = 'SAV case = {}'.format(os.path.basename(self.sav_case))

		# Update command and radio button status
		self.cmd_select_sav_case.config(text=lbl_sav_button)
		# self.year_om.config(state=Tk.NORMAL)
		# self.season_om.config(state=Tk.NORMAL)
		self.enable_radio_buttons(self.load_radio_btn_list)
		self.cmd_scale_load_gen.configure(state=Tk.NORMAL)

		return None

	def process(self):
		"""
			Function sorts the files list to remove any duplicates and then closes GUI window
		:return: None
		"""
		# Ask user to select target folder
		# target_file = tkFileDialog.asksaveasfilename(
		# 	initialdir=self.results_pth,
		# 	defaultextension='.xlsx',
		# 	filetypes=constants.General.file_types,
		# 	title='Please select file for results')
		#
		# if not target_file:
		# 	# Confirm a target file has actually been selected
		# 	_ = tkMessageBox.showerror(
		# 		title='No results file selected',
		# 		message='Please select a results file to save the fault currents to!'
		# 	)
		# elif not self.sav_case:
		# 	# Confirm SAV case has actually been provided
		# 	self.logger.error('No SAV case has been selected, please select SAV case')
		# 	# Display pop-up warning message
		# 	# Ask user to confirm that they actually want to close the window
		# 	_ = tkMessageBox.showerror(
		# 		title='No SAV case',
		# 		message='No SAV case has been selected for fault studies to be run on!'
		# 	)
		#
		# else:
		# 	# Save path to target file
		# 	self.target_file = target_file
		# 	# If SAV case has been selected the continue with study
		# 	# Process the fault times into useful format converting into floats
		# 	fault_times = self.var_fault_times_list.get()
		# 	fault_times = fault_times.split(',')
		# 	# Loop through each value converting to a float
		# 	# TODO: Add in more error processing here
		# 	self.fault_times = list()
		# 	for val in fault_times:
		# 		try:
		# 			new_val = float(val)
		# 			self.fault_times.append(new_val)
		# 		except ValueError:
		# 			self.logger.warning(
		# 				'Unable to convert the fault time <{}> to a number and so has been skipped'.format(val)
		# 			)
		#
		# 	self.logger.info(
		# 		(
		# 			'Faults will be applied at the busbars listed below and results saved to:\n{} \n '
		# 			'for the fault times: {} seconds.  \nBusbars = \n{}'
		# 		).format(self.target_file, self.fault_times, self.selected_busbars)
		# 	)

		print self.sav_case

		Load_Estimates_to_PSSE.update_loads(
			self.sav_case,
			constants.GUI.station_dict,
			year=self.year_selected.get(),
			season=self.season_selected.get()
		)

		_ = tkMessageBox.showinfo(
			title='Update Complete',
			message='PSSE sav case loads updated'
		)
		# Destroy GUI
		# self.master.destroy()
		return None

	def on_closing(self):
		"""
			Function runs when window is closed to determine if user actually wants to cancel running of study
		:return None:
		"""
		# Ask user to confirm that they actually want to close the window
		result = tkMessageBox.askquestion(
			title='Exit?',
			message='Are you sure you want to exit',
			icon='warning'
		)

		# Test what option the user provided
		if result == 'yes':
			# Close window
			self.master.destroy()
			self.abort = True
		else:
			return None


class BusbarsWindow:
	"""
		Produces a new window which is used to contain a list of busbars which can be edited by the user and
		then adjusted to produce the required output.
	"""
	def __init__(self, master, busbars=list()):
		"""
			Produce window which contains details of all the busbars that will be faulted and the ability to edit
			the busbars lists
		:param Tk.Tk() master:  This is the main window master which this one will be popped up on-top of
		:param list busbars:  List of busbars to be faulted will initially be populated into popup
		"""
		# Define initial values
		self.busbars = busbars
		self.entries = dict()

		# Get constants for how many busbars to display horizontally and vertically
		self.columns = constants.GUI.busbar_columns
		self.vertical_busbars = constants.GUI.vertical_busbars

		# Ensure that on_closing is processed correctly
		self.master = master
		self.master.protocol("WM_DELETE_WINDOW", self.on_closing)
		self.master.title = 'Identified Busbars'

		# Add label telling what is displayed in the box below
		lbl = Tk.Label(self.master, text='Edit busbar numbers to be faulted')
		lbl.grid(row=0, column=0)
		_ = CreateToolTip(lbl, text='Add or delete busbars to the list below and close window to continue.')

		# Produce a temporary frame which will house the entry boxes and place it in the window
		# TODO: May need to define height and width in real time
		busbar_frame = Tk.Frame(self.master, relief=Tk.GROOVE, bd=1)
		busbar_frame.grid(row=1, column=0)

		# Create a canvas into which the busbar_frame will be housed so that scroll bars can be added for the
		# busbar list
		self.canvas = Tk.Canvas(busbar_frame)
		# Add the canvas into a new frame
		self.entry_frame = Tk.Frame(self.canvas)
		# Create scroll bars which will control the busbar_frame within the canvas and configure the controls
		busbar_scrollbar = Tk.Scrollbar(busbar_frame, orient="vertical", command=self.canvas.yview)
		self.canvas.configure(yscrollcommand=busbar_scrollbar.set)

		# Locate the scroll bars on the right hand side of the canvas and locate canvas within newly created frame
		busbar_scrollbar.pack(side="right", fill="y")
		_ = CreateToolTip(busbar_scrollbar, text=(
			'If more busbars are needed then once out of space, close window and reopen to create additional space.'
		))
		self.canvas.pack(side="left")
		self.canvas.create_window((0, 0), window=self.entry_frame, anchor='nw')

		# Bind the action of the scrollbar with the movement of the canvas
		self.entry_frame.bind("<Configure>", self.canvas_scroll_function)

	def canvas_scroll_function(self, _event):
		"""
			Function to control what happens when the frame is scrolled
		:return:
		"""
		# Calculate the required width assuming busbar entry box size width is 60 pixels
		width = self.columns * 60
		# Calculate the required height assuming busbar entry box size height is 20 pixels
		height = self.vertical_busbars * 20
		self.canvas.configure(scrollregion=self.canvas.bbox("all"), width=width, height=height)

	def add_entry_boxes(self, spare_busbars=constants.GUI.empty_busbars):
		"""
			Adds the entry boxes ready to be populated with busbars
		:param int spare_busbars:  (optional=30) Number of extra rows to allow for additional busbars
		:return:
		"""
		# Calculate the number of entry boxes that are necessary allowing for extra number of busbars
		# Has to convert to float to ensure rounds up when dividing
		table_height = int(math.ceil(float(len(self.busbars)+spare_busbars) / float(self.columns)))
		counter = 0

		# Loop through each row and column adding a new entry box
		for row in xrange(table_height):
			for column in xrange(self.columns):
				# Create entry box for each busbar of the required size
				self.entries[counter] = Tk.Entry(self.entry_frame, width=constants.GUI.busbar_box_size)
				self.entries[counter].grid(row=row, column=column)
				counter += 1

		return None

	def populate_busbars(self):
		"""
			Function will replace all entries in the box with busbars provided on input
		:return None:
		"""
		# Populate entry box until all busbars reached
		for i, bus in enumerate(self.busbars):
			self.entries[i].insert(0, str(bus))

		return None

	def on_closing(self):
		"""
			Function will run when closed and update the list of busbars to be faulted
		:return None:
		"""
		busbars = list()
		for entry in self.entries.values():
			# Retrieve the value and continue to next entry box if empty
			value = entry.get()
			if value == '':
				continue

			# Try and convert the string input to an integer
			try:
				busbar = int(value)
				busbars.append(busbar)
			except ValueError:
				# There is an error with the value that has been provided, ask the user if they want to continue or
				# correct this number
				result = tkMessageBox.askquestion(
					title='Continue without busbar?',
					message=(
						'Unable to convert the busbar entry <{}> to a busbar number, do you want to ignore this value?'
					).format(entry.get()),
					icon='warning'
				)

				# Test what option the user provided
				if result == 'yes':
					# Skip this value
					continue
				else:
					# Give the user an option to correct
					print('Please correct the value {} to a busbar number'.format(value))
					return None

		# Identify busbars which either need to be removed or added to the list ensuring no duplicates
		busbars_to_remove = set(self.busbars)-set(busbars)
		busbars_to_add = set(busbars)-set(self.busbars)

		# Loop through each busbar and delete those which are no longer included
		for bus in busbars_to_remove:
			self.busbars.remove(bus)
		# Loop through each busbar and add those which are new
		for bus in busbars_to_add:
			self.busbars.append(bus)

		# Destroy popup window
		self.master.destroy()
		return None


class CreateToolTip(object):
	"""
		Function to create a popup tool tip for a given widget based on the descriptions provided here:
			https://stackoverflow.com/questions/3221956/how-do-i-display-tooltips-in-tkinter
	"""
	def __init__(self, widget, text='widget info'):
		"""
			Establish link with tooltip
		:param widget:  Tkinter element that tooltip should be associated with
		:param str text:  Message to display when hovering over button
		"""
		self.wait_time = 500     # milliseconds
		self.wrap_length = 450   # pixels
		self.widget = widget
		self.text = text
		self.widget.bind("<Enter>", self.enter)
		self.widget.bind("<Leave>", self.leave)
		self.widget.bind("<ButtonPress>", self.leave)
		self.id = None
		self.tw = None

	def enter(self, event=None):
		del event
		self.schedule()

	def leave(self, event=None):
		del event
		self.unschedule()
		self.hidetip()

	def schedule(self, event=None):
		del event
		self.unschedule()
		self.id = self.widget.after(self.wait_time, self.showtip)

	def unschedule(self, event=None):
		del event
		_id = self.id
		self.id = None
		if _id:
			self.widget.after_cancel(_id)

	def showtip(self):
		x, y, cx, cy = self.widget.bbox("insert")
		x += self.widget.winfo_rootx() + 25
		y += self.widget.winfo_rooty() + 20
		# creates a top level window
		self.tw = Tk.Toplevel(self.widget)
		self.tw.attributes('-topmost', 'true')
		# Leaves only the label and removes the app window
		self.tw.wm_overrideredirect(True)
		self.tw.wm_geometry("+%d+%d" % (x, y))
		label = Tk.Label(
			self.tw, text=self.text, justify='left', background="#ffffff", relief='solid', borderwidth=1,
			wraplength=self.wrap_length
		)
		label.pack(ipadx=1)

	def hidetip(self):
		tw = self.tw
		self.tw = None
		if tw:
			tw.destroy()


if __name__ == '__main__':
	gui = MainGUI()

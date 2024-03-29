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
import webbrowser
from PIL import Image, ImageTk
from collections import OrderedDict

# Package specific imports
import load_est
import load_est.psse as psse
import load_est.constants as constants
import load_est.dataframe_maker_modifier as dataframe_maker_modifier
import load_est.common_functions as common_functions
import load_est.scale as scale

# TODO: @NS - References to the Main script should be removed so that it can become a stand alone package
import Load_Estimates_to_PSSE


class CustomStyles:
    """ Class used to customize the layout of the GUI """

    def __init__(self):
        """
			Initialise the reference to the style names
		"""
        self.bg_color_frame = constants.GUIDefaultValues.color_frame
        self.bg_color_scrollbar = constants.GUIDefaultValues.color_scrollbar

        # Style for all other buttons

        # Style names
        # Buttons
        self.cmd_buttons = 'TButton'
        self.load_sav = 'LoadSav.TButton'
        self.cmd_run = 'Run.TButton'
        self.divers_button = 'Diverse.TButton'
        self.save_case_button = 'Save.TButton'
        self.update_case_button = 'Update.TButton'

        # Menus
        self.rating_options = 'TMenubutton'

        # ComboBox for selecting time
        self.combo_box = 'TCombobox'

        # Styles for labels
        self.label_general = 'TLabel'
        self.label_lbl_frame = 'LabelFrame.TLabel'
        self.label_mainheading = 'MainHeading.TLabel'
        self.label_subheading = 'SubHeading.TLabel'
        self.label_version_number = 'Version.TLabel'
        self.label_notes = 'Notes.TLabel'
        self.label_psc_info = 'PSCInfo.TLabel'
        self.label_psc_phone = 'PSCPhone.TLabel'
        self.label_hyperlink = 'Hyperlink.TLabel'

        # Radio buttons
        self.radio_buttons = 'TRadiobutton'

        # Check buttons
        self.check_buttons = 'TCheckbutton'
        self.check_buttons_framed = 'Framed.TCheckbutton'

        # Scroll bars
        self.scrollbar = 'TScrollbar'

        # Frames
        self.frame = 'TFrame'
        self.frame_outer = 'Outer.TFrame'
        self.lbl_frame = 'TLabelframe'

    def configure_styles(self):
        """
			Function configures all the ttk styles used within the GUI
			Further details here:  https://anzeljg.github.io/rin2/book2/2405/docs/tkinter/ttk-style-layer.html
		:return:
		"""
        # Tidy up the repeat ttk.Style() calls
        # Switch to a different theme
        styles = ttk.Style()
        styles.theme_use('winnative')

        # Configure the same font in all labels
        standard_font = constants.GUIDefaultValues.font_family
        bg_color = constants.GUIDefaultValues.color_main_window

        s = ttk.Style()
        s.configure('.', font=(standard_font, '8'))

        # General style for all buttons and active color changes
        s.configure(self.cmd_buttons, height=2, width=35)
        s.configure(self.load_sav, height=2, width=30)

        s.configure(self.cmd_run, height=2, width=50)

        s.configure(self.divers_button, height=2, width=20)
        s.configure(self.save_case_button, height=2, width=20)
        s.configure(self.update_case_button, height=2, width=20)

        s.configure(self.rating_options, height=2, width=15)

        s.configure(self.combo_box, height=2, width=50, state='readonly')

        s.configure(self.label_general, background=bg_color)
        s.configure(self.label_lbl_frame, font=(standard_font, '10'), background=bg_color, foreground='blue')
        s.configure(self.label_mainheading, font=(standard_font, '10', 'bold'), background=bg_color)
        s.configure(self.label_subheading, font=(standard_font, '9'), background=bg_color)
        s.configure(self.label_version_number, font=(standard_font, '7'), background=bg_color)
        s.configure(self.label_notes, font=(standard_font, '7'), background=bg_color)
        s.configure(self.label_hyperlink, foreground='Blue', font=(standard_font, '7'))

        s.configure(
            self.label_psc_info, font=constants.GUIDefaultValues.psc_font,
            color=constants.GUIDefaultValues.psc_color_web_blue, justify='center', background=bg_color
        )

        s.configure(
            self.label_psc_phone, font=(constants.GUIDefaultValues.psc_font, '8'),
            color=constants.GUIDefaultValues.psc_color_grey, background=bg_color
        )

        s.configure(self.radio_buttons, background=bg_color)
        s.map(self.radio_buttons, background=[('disabled', bg_color)])

        # Ensure style for check buttons is tick rather than cross
        s.configure(self.check_buttons, background=bg_color)
        s.configure(self.check_buttons_framed, background=self.bg_color_frame)

        # Configure the scroll bar
        s.configure(self.scrollbar, background=self.bg_color_scrollbar)

        # Frames used to house other drawings
        s.configure(self.frame, background=self.bg_color_frame)
        s.configure(self.frame_outer, bd=1, background=self.bg_color_frame)

        # configure label frame
        s.configure(self.lbl_frame, font=('courier', 15, 'bold'), foreground='blue', background=bg_color)


def enable_radio_buttons(radio_btn_list, enable=True):
    """
		Function to enable or disable radio buttons that are stored in radio_btn_list
	:param radio_btn_list: list of radio button widgets
	:param enable: True or False
	:return:
	"""

    if enable:
        for radio_btn in radio_btn_list:
            radio_btn.config(state=Tk.NORMAL)
    else:
        for radio_btn in radio_btn_list:
            radio_btn.config(state=Tk.DISABLED)

    return None


class MainGUI:
    """
		Main class to produce the GUI
		Allows the user to select the busbars and methodology to be applied in the fault current calculations
	"""
    # General constants which need to be initialised
    _row = 0
    _col = 0
    xpad = 5
    ypad = 5

    def __init__(self, title=constants.GUI.gui_name):
        """
			Initialise GUI
		:param str title: (optional) - Title to be used for main window
		"""

        # Get logger handle
        self.logger = logging.getLogger(constants.Logging.logger_name)
        self.abort = False

        # Initialise constants and Tk window
        self.master = Tk.Tk()
        self.master.title(title)

        # Change color of main window
        self.master.configure(bg=constants.GUIDefaultValues.color_main_window)

        # Ensure that on_closing is processed correctly
        self.master.protocol("WM_DELETE_WINDOW", self.on_closing)

        # Get reference to the custom styles being used
        self.styles = CustomStyles()
        self.styles.configure_styles()

        # Target file that results will be exported to
        self.target_file = str()
        self.results_pth = os.path.dirname(os.path.realpath(__file__))

        # Stand alone command buttons
        self.cmd_select_sav_case = ttk.Button()
        self.sav_case_status = Tk.NORMAL
        self.cmd_import_load_estimates = ttk.Button()
        self.cmd_scale_load_gen = ttk.Button()

        # PSC logo constants (required so position can be moved)
        self.psc_info = Tk.Label()
        self.hyp_user_manual = ttk.Label()
        self.version_tool_lbl = ttk.Label()

        # excel and sav case labels
        self.current_xl_path_lbl = ttk.Label()
        self.load_complete_lbl_t_f = ttk.Label()

        # PSSE variables
        self.sav_case = str()
        self.psse_con = None

        # ---------------------------------------- LOAD SCALING OPTIONS:------------------------------------------------
        # Load Options label frame constants
        self.load_labelframe = ttk.LabelFrame()
        # boolean variable to store which scrollable frame the mouse is in
        self._inLoadFrame = bool()

        # Load radio buttons variables
        self.load_radio_opt_sel = Tk.IntVar()
        self.load_radio_btn_list = list()
        self.load_radio_opts = OrderedDict()
        self.load_prev_radio_opt = int()

        # year drop down options
        self.load_year_selected = Tk.StringVar()
        self.load_year_list = ['N/A']
        self.load_year_om = None
        self.load_year_om_lbl = ttk.Label()

        # season drop down options
        self.load_demand_scaling_selected = Tk.StringVar()
        self.load_demand_scaling_list = common_functions.Headers.seasons
        self.load_demand_scaling_om = None
        self.load_demand_scaling_om_lbl = ttk.Label()

        # divers load drop down options
        self.load_divers_selected = Tk.StringVar()
        self.load_divers_list = common_functions.Headers.loads_divers
        self.load_divers_om = None
        self.load_divers_om_lbl = ttk.Label()

        # Load Selector scroll constants
        self.load_side_lbl = ttk.Label()
        self.load_select_frame = None
        self.load_canvas = Tk.Canvas()
        self.load_entry_frame = ttk.Frame()
        self.load_scrollbar = ttk.Scrollbar()
        # Load selector zone constants
        self.load_bool_var_zone = dict()
        self.load_zones_selected = dict()

        self.load_bool_var_gsp = dict()

        self.load_gsps = ['N/A']
        self.load_gsps_selected = list()

        # ---------------------------------------- GENERATION SCALING OPTIONS:------------------------------------------
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
        self.gen_bool_var_zone = dict()
        self.gen_zones_selected = dict()

        self.gen_percent_var = Tk.DoubleVar()
        # Add entry box
        self.gen_percent_entry = Tk.Entry()

        # Zone data variables
        self.zone_data = None
        self.zone_dict = dict()

        # update GUI variables to create the GUI
        self.create_gui()

        self.logger.debug('GUI window created')
        # Produce GUI window
        self.master.mainloop()

    def create_gui(self):
        """
		Function to update initialise GUI variables to create the GUI
		:return:
		"""
        # Add PSC logo with hyperlink to the website
        self.add_psc_logo(row=self.row(), col=0)

        # add import load estimate excel button
        self.cmd_import_load_estimates = self.add_cmd_button(
            label='Import SSE Load Estimates .xlsx', cmd=self.import_load_estimates_xl,
            row=self.row(), col=3)

        # add excel summary labels:
        self.add_xl_summary_labels()

        # Add PSC logo with hyperlink to the website
        self.add_psc_logo(row=0, col=5)

        # Determine label for button based on the SAV case being loaded
        if self.sav_case:
            lbl_sav_button = 'SAV case = {}'.format(os.path.basename(self.sav_case))
        else:
            lbl_sav_button = 'Select PSSE SAV Case'

        self.cmd_select_sav_case = self.add_cmd_button(
            label=lbl_sav_button, cmd=self.select_sav_case,
            row=self.row(1), col=3)
        self.cmd_select_sav_case.configure(state=self.sav_case_status)

        self.add_sep(row=4, col_span=8)

        # create load options labelframe
        self.create_load_options()

        # create generation options labelframe
        self.create_gen_options()

        # set ._col and ._row to the size of grid
        self._col, self._row, = self.master.grid_size()
        # add cmd button to scale load
        self.cmd_scale_load_gen = self.add_cmd_button(
            label='Scale Load and Generation', cmd=self.scale_loads_gens,
            row=self.row(1), col=3)
        self.cmd_scale_load_gen.configure(state=Tk.DISABLED)

        # Add PSC logo in Windows Manager
        self.add_psc_logo_wm()

        # set ._col and ._row to the size of grid
        self._col, self._row, = self.master.grid_size()
        # Add PSC UK and phone number
        self.add_psc_phone(row=self.row(), col=0)
        self.add_hyp_user_manual(row=self.row(), col=self.col() - 1)

        return

    def add_xl_summary_labels(self):
        """
		Function to add summary labels about the loaded excel file
		:return:
		"""
        # add Current File Loaded label
        current_xl_lbl = ttk.Label(self.master, text='Current File Loaded:', style=self.styles.label_general)
        current_xl_lbl.grid(row=self.row(1), column=3)

        if constants.General.xl_file_name:
            xl_lbl = constants.General.xl_file_name
            self.sav_case_status = Tk.NORMAL

            if constants.General.loads_complete:
                com_lbl = constants.General.loads_complete_t_str
                lbl_color = 'green'
            else:
                com_lbl = constants.General.loads_complete_f_str
                lbl_color = 'red'
        else:
            xl_lbl = 'N/A'
            com_lbl = 'N/A'
            lbl_color = 'black'
            self.sav_case_status = Tk.DISABLED

        # add excel path label
        self.current_xl_path_lbl = ttk.Label(self.master, text=xl_lbl, style=self.styles.label_general)
        self.current_xl_path_lbl.grid(row=self.row(), column=4)

        # add load complete label path label
        load_complete_lbl = ttk.Label(self.master, text='All loads error free:', style=self.styles.label_general)
        load_complete_lbl.grid(row=self.row(1), column=3)

        # add True/False label
        self.load_complete_lbl_t_f = ttk.Label(
            self.master, text=str(com_lbl), style=self.styles.label_general, foreground=lbl_color, cursor='')
        self.load_complete_lbl_t_f.grid(row=self.row(), column=4)

        # if there are errors in the load file make the label open the excel summary file
        if not constants.General.loads_complete and constants.General.xl_file_name:
            # file_path = os.path.join(
            #     constants.General.curPath,
            #     constants.XlFileConstants.params_folder,
            #     constants.XlFileConstants.xl_checks_file_name)
            file_path = common_functions.get_local_file_path(file_name=constants.General.xl_file_name)
            # todo:@Ask_David: the following line does not open the excel file
            self.load_complete_lbl_t_f.bind("<Button-1>", lambda e: webbrowser.open_new(file_path))
            self.load_complete_lbl_t_f.configure(cursor='hand2')

        return

    def create_load_options(self):
        """
		Function to create the load scaling options for selection of different years
		:return:
		"""
        # add Load Scaling Options: label
        label = ttk.Label(self.master, text='Load Scaling Options:', style=self.styles.label_lbl_frame)
        self.load_labelframe = ttk.LabelFrame(self.master, labelwidget=label, style=self.styles.lbl_frame)
        self.load_labelframe.grid(row=5, column=0, columnspan=4, padx=self.xpad, pady=self.ypad, sticky='NSEW')

        # Initial option 0 as the first option and set as previous option also
        self.load_radio_opt_sel = Tk.IntVar(self.master, 0)
        self.load_prev_radio_opt = 0

        # Dictionary to create multiple buttons
        self.load_radio_opts = constants.GUI.load_radio_opts

        self._row = 0
        self.load_radio_btn_list = list()
        # Loop is used to create multiple Radio buttons,  rather than creating each button separately
        for (value, text) in self.load_radio_opts.items():
            temp_rb = ttk.Radiobutton(
                self.load_labelframe,
                text=text,
                variable=self.load_radio_opt_sel,
                value=value,
                command=self.load_radio_button_click,
                style=self.styles.radio_buttons
            )
            temp_rb.grid(row=self.row(), columnspan=2, sticky='W', padx=self.xpad, pady=self.ypad)
            self.load_radio_btn_list.append(temp_rb)
            self.row(1)

        # disable all radio buttons in load_radio_btn_list
        enable_radio_buttons(self.load_radio_btn_list, enable=False)

        # add drop down list for year
        self.load_year_selected = Tk.StringVar(self.master)

        # add label for year drop down box, # grey out text initially
        self.load_year_om_lbl = ttk.Label(master=self.load_labelframe, text='Year: ', style=self.styles.label_general)
        self.load_year_om_lbl.grid(row=self.row(), column=self.col(), rowspan=1, sticky=Tk.E, padx=self.xpad)
        self.load_year_om_lbl.configure(foreground='grey')

        # add year drop down box, initially disabled
        self.load_year_om = self.add_combobox(
            row=self.row(),
            col=self.col(1),
            var=self.load_year_selected,
            values_list=self.load_year_list,
            location=self.load_labelframe
        )
        self.load_year_om.configure(state=Tk.DISABLED)

        self._col = 0
        # add drop down list for season
        self.load_demand_scaling_selected = Tk.StringVar(self.master)

        # add demand scaling label, grey out text initially
        self.load_demand_scaling_om_lbl = ttk.Label(
            master=self.load_labelframe, text='Demand Scaling: ', style=self.styles.label_general)
        self.load_demand_scaling_om_lbl.grid(row=self.row(1), column=self.col(), rowspan=1, sticky=Tk.E, padx=self.xpad)
        self.load_demand_scaling_om_lbl.configure(foreground='grey')

        # add demand scaling label, initially disabled
        self.load_demand_scaling_om = self.add_combobox(
            row=self.row(), col=self.col(1),
            var=self.load_demand_scaling_selected,
            values_list=self.load_demand_scaling_list,
            location=self.load_labelframe
        )
        self.load_demand_scaling_om.configure(state=Tk.DISABLED)

        self._col = 0
        # add drop down list for divers
        self.load_divers_selected = Tk.StringVar(self.master)

        # add demand scaling label, grey out text initially
        self.load_divers_om_lbl = ttk.Label(
            master=self.load_labelframe, text='Diverse/Aggregate: ', style=self.styles.label_general)
        self.load_divers_om_lbl.grid(row=self.row(2), column=self.col(), rowspan=1, sticky=Tk.E, padx=self.xpad)
        self.load_divers_om_lbl.configure(foreground='grey')

        # add demand scaling label, initially disabled
        self.load_divers_om = self.add_combobox(
            row=self.row(), col=self.col(1),
            var=self.load_divers_selected,
            values_list=self.load_divers_list,
            location=self.load_labelframe
        )
        self.load_divers_om.configure(state=Tk.DISABLED)

        return None

    def create_gen_options(self):

        # add Generation Scaling Options: label
        label = ttk.Label(self.master, text='Generation Scaling Options:', style=self.styles.label_lbl_frame)
        self.gen_labelframe = ttk.LabelFrame(self.master, labelwidget=label, style=self.styles.lbl_frame)
        self.gen_labelframe.grid(row=5, column=4, columnspan=4, padx=self.xpad, pady=self.ypad, sticky='NSEW')

        # Initial option 0 as the first option and set as previous option also
        self.gen_radio_opt_sel = Tk.IntVar(self.master, 0)
        self.gen_prev_radio_opt = 0

        # Dictionary to create multiple buttons
        self.gen_radio_opts = constants.GUI.gen_radio_opts

        self._row = 0
        self.gen_radio_btn_list = list()
        # Loop is used to create multiple Radio buttons, rather than creating each button separately
        for (value, text) in self.gen_radio_opts.items():
            temp_rb = ttk.Radiobutton(
                self.gen_labelframe,
                text=text,
                variable=self.gen_radio_opt_sel,
                value=value,
                command=self.gen_radio_button_click,
                style=self.styles.radio_buttons
            )
            temp_rb.grid(row=self.row(), sticky='W', padx=self.xpad, pady=self.ypad)
            self.gen_radio_btn_list.append(temp_rb)
            self.row(1)

        # disable all radio buttons in gen_radio_btn_list
        enable_radio_buttons(self.gen_radio_btn_list, enable=False)

        # add entry box for generator scaling percentage
        self.add_entry_gen_percent(row=self.row(), col=0)

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

    def add_combobox(self, row, col, var, values_list, location=None):
        """
			Function to all a list of optimisation options which will become enabled if the user selects to run a
			virtual statcom study
		:param int row: Row number to use
		:param int col: Column number to use
		:param Tk.StringVar var:  Variable to link with combobox
		:param list values_list:  List of values to use as an input
		:param Tk.Labelframe location:  Parent Tk window in which to create drop down list
		:rtype Tk.OptionMenu
		:return dropdown_optimisation_option:
		"""

        # Check whether there is a successfully loaded SAV case to enable the list option
        if location is None:
            var.set(values_list[0])
            # Create the drop down list to be shown in the GUI
            w = ttk.Combobox(
                self.master,
                textvariable=var,
                values=values_list,
                style=self.styles.rating_options,
                justify=Tk.CENTER
            )
            w.grid(
                row=row, column=col, columnspan=1,
                padx=self.xpad, pady=self.ypad, sticky=Tk.W + Tk.E,
            )
        else:
            var.set(values_list[0])
            # Create the drop down list to be shown in the GUI
            # w = ttk.Combobox(location, var, list[0], *list, style=self.styles.rating_options)
            w = ttk.Combobox(
                location,
                textvariable=var,
                values=values_list,
                style=self.styles.rating_options,
                justify=Tk.CENTER,
            )

            w.grid(
                row=row, column=col, columnspan=1,
                padx=self.xpad, pady=self.ypad, sticky=Tk.W + Tk.E,
            )

        return w

    def add_cmd_button(self, label, cmd, row, col, location=None):
        """
			Function just adds the command button to the GUI which is used for selecting the SAV case
		:param int row: Row number to use
		:param int col: Column number to use
		:param str label:  Label to use
		:param cmd:  Function to call when button clicked
		:param Tk.Frame location:  Parent window to use
		:return None:
		"""
        if location is None:
            # Create button and assign to Grid
            cmd_btn = ttk.Button(self.master, text=label, command=cmd, style=self.styles.cmd_buttons)
            cmd_btn.grid(row=row, column=col, columnspan=2, sticky=Tk.W + Tk.E, padx=self.xpad, pady=self.ypad)
        # CreateToolTip(widget=self.cmd_select_sav_case, text=(
        # 	'Select the SAV case for which fault studies should be run.'
        # ))
        else:
            cmd_btn = ttk.Button(location, text=label, command=cmd, style=self.styles.cmd_buttons)
            # self.cmd_select_sav_case.grid(row=row, column=col, columnspan=2, sticky=Tk.W + Tk.E)
            cmd_btn.grid(row=row, column=col, columnspan=2, sticky=Tk.W + Tk.E, padx=self.xpad, pady=self.ypad)
        # CreateToolTip(widget=self.cmd_select_sav_case, text=(
        # 	'Select the SAV case for which fault studies should be run.'
        # ))

        return cmd_btn

    def add_hyp_user_manual(self, row, col):
        """
			Function just adds the version and hyperlink to the user manual to the GUI
		:param row: Row Number to use
		:param col: Column number to use
		:return None:
		"""
        # Create user manual link and reference to the version of the tool
        self.hyp_user_manual = ttk.Label(
            self.master, text="User Guide", cursor="hand2", style=self.styles.label_hyperlink)
        self.hyp_user_manual.grid(row=row, column=col, sticky="se", padx=self.xpad)
        self.hyp_user_manual.bind('<Button - 1>', lambda e: webbrowser.open_new(
            os.path.join(constants.GUIDefaultValues.local_directory, constants.GUIDefaultValues.user_guide)))
        self.version_tool_lbl = ttk.Label(
            self.master, text='Version {}'.format(constants.__version__), style=self.styles.label_version_number)
        self.version_tool_lbl.grid(row=row, column=col - 1, sticky="se", padx=self.xpad)
        CreateToolTip(widget=self.hyp_user_manual, text=(
            "Open the GUI user guide"
        ))
        CreateToolTip(widget=self.version_tool_lbl, text=(
            "Version of the tool"
        ))
        return None

    def scale_loads_gens(self):

        variable_name = [common_functions.folder_file_names.dill_good_data_name]
        dill_folder_name = common_functions.folder_file_names.dill_folder
        df = common_functions.load_dill(variable_name=variable_name, dill_folder_name=dill_folder_name)  # a good_data
        # dataframe is loaded here
        # f_load_values, year = str(), season = str(), diverse = False, zone = tuple(), gsp = tuple()

        # if self.load_radio_opt_sel.get() == 1:
        if self.load_divers_selected.get() == 'Diverse':
            diverse = True
        else:
            diverse = False

        # all loads
        if self.load_radio_opt_sel.get() == 1:
            # todo: @ask David, how to add a toggle for divers, gsp, and zones to give them as an input here
            # zone = ()
            # gsp = ()
            # diverse = True
            # zones = ()
            # gsps = ()
            # todo: this should be coming from a toggle of the GUI in load scale options

            # self.load_gsps_selected = list()
            # self.load_zones_selected = dict()
            # gsps = self.load_gsps_selected.get()
            # zones = self.load_zones_selected.get()

            # Load_Estimates_to_PSSE.scale_all_loads(
            # 	year=self.load_year_selected.get(),
            # 	season=self.load_demand_scaling_selected.get()
            # 	)
            scale.scale_loads(df_load_values=df, year=self.load_year_selected.get(),
                              season=self.load_demand_scaling_selected.get(),
                              diverse=diverse, zone=(),
                              gsp=())

            # scale.scale_loads(df_load_values=df,year=self.load_year_selected.get(),
            # season=self.load_demand_scaling_selected.get(), diverse=diverse ,zone=zone, gsp=gsp)
            # todo: @ask David:
            #  the following gsp and zone values are not obtained from the gui values, first the gui needs to update in
            #  set_gui_param
            # Check if All Loads selected
            k=1

            # Otherwise populate with selected GSPs / Zones
            # self.load_gsps_selected = (
            #     gsp_name for gsp_name, bool_var in self.load_bool_var_gsp.items() if bool_var.get() == 1
            # )
            # # todo: convert the diverse button value which is a string (diverese,aggeragte) to 0 or 1
            # scale.scale_loads(df_load_values=df, year=self.load_year_selected.get(),
            #                   season=self.load_demand_scaling_selected.get(),
            #                   diverse=self.load_divers_selected.get(), zone=self.load_zones_selected.get(),
            #                   gsp=self.load_gsps_selected.get())
        # gsps selected
        if self.load_radio_opt_sel.get() == 2:

            self.load_gsps_selected = (
                gsp_name for gsp_name, bool_var in self.load_bool_var_gsp.items() if bool_var.get() == 1
            )
            scale.scale_loads(df_load_values=df, year=self.load_year_selected.get(),
                              season=self.load_demand_scaling_selected.get(),
                              diverse=diverse, zone=(),
                              gsp=self.load_gsps_selected.get())
        # zones selected
        if self.load_radio_opt_sel.get() == 3:
            self.load_zones_selected = (
                zone_name for zone_name, bool_var in self.load_bool_var_zone.items() if bool_var.get() == 1
            )
            scale.scale_loads(df_load_values=df, year=self.load_year_selected.get(),
                              season=self.load_demand_scaling_selected.get(),
                              diverse=diverse, zone=self.load_zones_selected.get(),
                              gsp=())

        # all gens
        if self.gen_radio_opt_sel.get() == 1:
            # Load_Estimates_to_PSSE.scale_all_gens(
            # 	pc=self.gen_percent_entry.get()
            # 	)

            # scale.scale_gens(pc=self.gen_percent_entry.get())
            scale.scale_gens(pc=self.gen_percent_entry.get(), zone=())
        # selected zone gens
        if self.gen_radio_opt_sel.get() == 2:
            # Load_Estimates_to_PSSE.scale_all_gens(
            # 	pc=self.gen_percent_entry.get()
            # 	)
            self.gen_zones_selected = (
                zone_name for zone_name, bool_var in self.gen_bool_var_zone.items() if bool_var.get() == 1
            )
            # scale.scale_gens(pc=self.gen_percent_entry.get())
            scale.scale_gens(pc=self.gen_percent_entry.get(), zone=self.gen_zones_selected.get())

        # Reset to empty entries
        self.load_gsps_selected = list()
        self.load_zones_selected = dict()
        self.gen_zones_selected = dict()

        self.psse_con.update_psse_gui()

        return

    def load_radio_button_click(self):
        """
		Function that is called when a load radio button is clicked to enable buttons and create/remove load selection
		frame
		:return:
		"""
        # none load scaling option
        if self.load_radio_opt_sel.get() == 0:
            # disable year and seasons drop downs and labels
            self.load_year_om.config(state=Tk.DISABLED)
            self.load_year_om_lbl.config(foreground='grey')
            self.load_demand_scaling_om.config(state=Tk.DISABLED)
            self.load_demand_scaling_om_lbl.config(foreground='grey')
            self.load_divers_om_lbl.config(foreground='grey')
            self.load_divers_om.config(state=Tk.DISABLED)
            # remove the bottom load selector
            self.remove_load_scroll_frame()
        # All load scaling option
        elif self.load_radio_opt_sel.get() == 1:
            # enable year and seasons drop downs and labels
            self.load_year_om.config(state=Tk.NORMAL)
            self.load_year_om_lbl.config(foreground='')
            self.load_demand_scaling_om.config(state=Tk.NORMAL)
            self.load_demand_scaling_om_lbl.config(foreground='')
            self.load_divers_om_lbl.config(foreground='')
            self.load_divers_om.config(state=Tk.NORMAL)
            # remove the bottom load selector
            self.remove_load_scroll_frame()
            self.cmd_scale_load_gen.configure(state=Tk.NORMAL)
        # All load scaling option, for GSP/Zones selection
        elif self.load_radio_opt_sel.get() > 1 and self.load_radio_opt_sel.get() != self.load_prev_radio_opt:
            # enable year and seasons drop downs and labels
            self.load_year_om.config(state=Tk.NORMAL)
            self.load_year_om_lbl.config(foreground='')
            self.load_demand_scaling_om.config(state=Tk.NORMAL)
            self.load_demand_scaling_om_lbl.config(foreground='')
            self.load_divers_om_lbl.config(foreground='')
            self.load_divers_om.config(state=Tk.NORMAL)
            # create load selector frame
            self.create_load_select_frame()
            self.cmd_scale_load_gen.configure(state=Tk.DISABLED)

        # store last radio button selected in load_prev_radio_opt
        self.load_prev_radio_opt = self.load_radio_opt_sel.get()

        return None

    def gen_radio_button_click(self):
        """
		Function to control GUI actions when a generation radio button is clicked
		:return:
		"""

        if self.gen_radio_opt_sel.get() == 0:
            # disable entry box and remove the generation scroll frame
            self.gen_percent_entry.config(state=Tk.DISABLED)
            self.remove_gen_scroll_frame()

        elif self.gen_radio_opt_sel.get() == 1:
            # enable entry box and remove the generation scroll frame
            self.gen_percent_entry.config(state=Tk.NORMAL)
            self.remove_gen_scroll_frame()
            self.cmd_scale_load_gen.configure(state=Tk.NORMAL)

        elif self.gen_radio_opt_sel.get() > 1 and self.gen_radio_opt_sel.get() != self.gen_prev_radio_opt:
            # enable entry box and create the generation scroll frame
            self.gen_percent_entry.config(state=Tk.NORMAL)
            self.create_gen_select_frame()
            self.cmd_scale_load_gen.configure(state=Tk.DISABLED)

        # store last radio button selected in gen_prev_radio_opt
        self.gen_prev_radio_opt = self.gen_radio_opt_sel.get()

        return None

    def create_load_select_frame(self):
        # it creates a frame of GSP/zone user selection form
        lbl = ''
        if self.load_radio_opt_sel.get() == 2:
            lbl = 'Select GSP(s):'
        elif self.load_radio_opt_sel.get() == 3:
            lbl = 'Select Zone(s):'

        if self.load_select_frame is None:
            master_col, master_rows, = self.master.grid_size()
            self._row = master_rows
            self._col = 0
            # Label for what is included in entry
            self.load_side_lbl = ttk.Label(master=self.load_labelframe, text=lbl, style=self.styles.label_general)
            self.load_side_lbl.grid(row=self.row(1), column=self.col(), sticky=Tk.W, padx=self.xpad, pady=self.ypad)

            # Produce a frame which will house the zones check buttons and place them in the window
            self.load_select_frame = ttk.Frame(self.load_labelframe, relief=Tk.GROOVE, style=self.styles.frame_outer)
            self.load_select_frame.grid(
                row=self.row(1), column=self.col(), columnspan=4, sticky='NSEW', padx=self.xpad, pady=self.ypad)

            # Create a canvas into which the load_select_frame will be housed so that scroll bars can be added
            self.load_canvas = Tk.Canvas(self.load_select_frame)

            # Add the canvas into a new frame
            self.load_entry_frame = ttk.Frame(self.load_canvas, style=self.styles.frame)

            # Create scroll bars which will control the zon_frame within the canvas and configure the controls
            self.load_scrollbar = ttk.Scrollbar(
                self.load_select_frame, orient="vertical", command=self.load_canvas.yview, style=self.styles.scrollbar
            )
            self.load_canvas.configure(yscrollcommand=self.load_scrollbar.set)

            # Locate the scroll bars on the right hand side of the canvas and locate canvas within newly created frame
            self.load_scrollbar.pack(side="right", fill="y")
            self.load_canvas.pack(side="left")
            self.load_canvas.create_window((0, 0), window=self.load_entry_frame, anchor='nw')

            # Bind the action of the scrollbar with the movement of the canvas
            self.load_entry_frame.bind("<Configure>", self.load_canvas_scroll_function)
            self.load_select_frame.bind("<Enter>", self._bound_to_mousewheel)
            self.load_select_frame.bind("<Leave>", self._unbound_to_mousewheel)
            self.move_widgets_down()

        else:
            self.load_side_lbl.config(text=lbl)
            if len(self.load_entry_frame.winfo_children()) > 0:
                for widget in self.load_entry_frame.winfo_children():
                    widget.destroy()
            self.load_bool_var_gsp = dict()
            self.load_bool_var_zone = dict()
            self.gen_bool_var_zone = dict()

        if self.load_radio_opt_sel.get() == 2:
            # 'Select GSP(s):'
            for gsp in self.load_gsps:
                self.load_bool_var_gsp[gsp] = Tk.BooleanVar()
                self.load_bool_var_gsp[gsp].set(0)

            counter = 0
            for gsp in self.load_gsps:
                # lbl = "{}".format(self.load_gsps[gsp])
                lbl = "{}".format(gsp)
                check_button = ttk.Checkbutton(
                    self.load_entry_frame,
                    text=lbl,
                    variable=self.load_bool_var_gsp[gsp],
                    style=self.styles.check_buttons_framed
                )
                check_button.grid(row=counter, column=0, sticky="w")
                counter += 1
            self.load_canvas.yview_moveto(0)

        if self.load_radio_opt_sel.get() == 3:
            # 'Select Zone(s):'
            for zone_num, zone_name in self.zone_dict.iteritems():
                self.load_bool_var_zone[zone_num] = Tk.BooleanVar()
                self.load_bool_var_zone[zone_num].set(0)

            counter = 0
            for zone_num, zone_name in self.zone_dict.iteritems():
                lbl = "{}".format(str(zone_num) + '. ' + str(zone_name))
                check_button = ttk.Checkbutton(
                    self.load_entry_frame,
                    text=lbl,
                    variable=self.load_bool_var_zone[zone_num],
                    style=self.styles.check_buttons_framed
                )
                check_button.grid(row=counter, column=0, sticky="w")
                counter += 1
            self.load_canvas.yview_moveto(0)

        return None

    def create_gen_select_frame(self):
        """
			Create a frame for selection of generation data
		"""

        if self.gen_radio_opt_sel.get() == 2:
            lbl = 'Select Zone(s):'
        else:
            lbl = str()

        if self.gen_select_frame is None:
            master_col, master_rows, = self.master.grid_size()
            self._row = master_rows
            self._col = 0
            # Label for what is included in entry
            self.gen_side_lbl = ttk.Label(master=self.gen_labelframe, text=lbl, style=self.styles.label_general)
            self.gen_side_lbl.grid(row=self.row(1), column=self.col(), sticky='W', padx=self.xpad, pady=self.ypad)

            # Produce a frame which will house the zones check buttons and place them in the window
            self.gen_select_frame = ttk.Frame(self.gen_labelframe, relief=Tk.GROOVE, style=self.styles.frame_outer)
            self.gen_select_frame.grid(
                row=self.row(1), column=self.col(), columnspan=4, sticky='NSEW', padx=self.xpad, pady=self.ypad)

            # Create a canvas into which the zon_frame will be housed so that scroll bars can be added for the
            # network zone list
            self.gen_canvas = Tk.Canvas(self.gen_select_frame)

            # Add the canvas into a new frame
            self.gen_entry_frame = ttk.Frame(self.gen_canvas, style=self.styles.frame)

            # Create scroll bars which will control the zon_frame within the canvas and configure the controls
            self.gen_scrollbar = ttk.Scrollbar(
                self.gen_select_frame, orient="vertical", command=self.gen_canvas.yview, style=self.styles.scrollbar
            )
            self.gen_canvas.configure(yscrollcommand=self.gen_scrollbar.set)

            # Locate the scroll bars on the right hand side of the canvas and locate canvas within newly created frame
            self.gen_scrollbar.pack(side="right", fill="y")
            self.gen_canvas.pack(side="left")
            self.gen_canvas.create_window((0, 0), window=self.gen_entry_frame, anchor='nw')

            # Bind the action of the scrollbar with the movement of the canvas
            self.gen_entry_frame.bind("<Configure>", self.gen_canvas_scroll_function)
            self.gen_select_frame.bind("<Enter>", self._bound_to_mousewheel)
            self.gen_select_frame.bind("<Leave>", self._unbound_to_mousewheel)
            self.move_widgets_down()

        else:
            if len(self.gen_entry_frame.winfo_children()) > 0:
                for widget in self.gen_entry_frame.winfo_children():
                    widget.destroy()
            self.gen_zones_selected = dict()

        if self.gen_radio_opt_sel.get() == 2:
            lbl = 'Select Zone(s):'
            self.gen_side_lbl.config(text=lbl)
            # 'Select Zone(s):'
            for zone_num, zone_name in self.zone_dict.iteritems():
                self.gen_bool_var_zone[zone_num] = Tk.BooleanVar()
                self.gen_bool_var_zone[zone_num].set(0)

            counter = 0
            for zone_num, zone_name in self.zone_dict.iteritems():
                lbl = "{}".format(str(zone_num) + '. ' + str(zone_name))
                check_button = ttk.Checkbutton(
                    self.gen_entry_frame,
                    text=lbl,
                    variable=self.gen_bool_var_zone[zone_num],
                    style=self.styles.check_buttons_framed
                )
                check_button.grid(row=counter, column=0, sticky="w")
                counter += 1
            self.gen_canvas.yview_moveto(0)

        return None

    def _bound_to_mousewheel(self, event):
        # Function to bind scrolling in the load or generation scroll frame
        if event.widget == self.load_select_frame:
            self.load_select_frame.bind_all("<MouseWheel>", self._on_mousewheel)
            self._inLoadFrame = True
        if event.widget == self.gen_select_frame:
            self.gen_select_frame.bind_all("<MouseWheel>", self._on_mousewheel)
            self._inLoadFrame = False

    def _unbound_to_mousewheel(self, event):
        # Function to unbind scrolling in the load or generation scroll frame
        if event.widget == self.load_select_frame:
            self.load_select_frame.unbind_all("<MouseWheel>")
            self._inLoadFrame = bool()
        if event.widget == self.gen_select_frame:
            self.gen_select_frame.unbind_all("<MouseWheel>")
            self._inLoadFrame = bool()

    def _on_mousewheel(self, event):
        # Function to scroll the load or generation scroll frame depend on which
        if self._inLoadFrame:
            self.load_canvas.yview_scroll(-1 * (event.delta / 120), "units")
        else:
            self.gen_canvas.yview_scroll(-1 * (event.delta / 120), "units")

    def load_canvas_scroll_function(self, _event):
        # Function to set the scrolling region in the load frame
        self.load_canvas.configure(scrollregion=self.load_canvas.bbox("all"), width=230, height=200)
        return None

    def gen_canvas_scroll_function(self, _event):
        # Function to set the scrolling region in the load frame
        self.gen_canvas.configure(scrollregion=self.gen_canvas.bbox("all"), width=230, height=200)
        return None

    def remove_load_scroll_frame(self):
        """
		Function to remove load scroll frame by removing widgets
		:return:
		"""

        if self.load_select_frame is not None:
            # remove widgets
            self.load_side_lbl.grid_remove()
            self.load_select_frame.grid_remove()

            # reset variables
            self.load_side_lbl = None
            self.load_select_frame = None
            self.load_bool_var_gsp = dict()
            self.load_bool_var_zone = dict()
            self.gen_bool_var_zone = dict()

        return None

    def remove_gen_scroll_frame(self):
        """
		Function to remove generation scroll frame by removing widgets/
		:return:
		"""

        if self.gen_select_frame is not None:
            # remove widgets
            self.gen_side_lbl.grid_remove()
            self.gen_select_frame.grid_remove()

            # reset variables
            self.gen_side_lbl = None
            self.gen_select_frame = None
            self.load_bool_var_gsp = dict()
            self.load_bool_var_zone = dict()
            self.gen_bool_var_zone = dict()

        return None

    def move_widgets_down(self):
        """
		Function called when a scroll frame is created to moved widget down accordingly
		:return:
		"""

        # find the end column and end row of the current grid
        master_col, master_rows, = self.master.grid_size()

        # change grid location of widgets
        self.cmd_scale_load_gen.grid(
            row=master_rows + 1, column=self.cmd_scale_load_gen.grid_info()['column'])
        self.psc_info.grid(
            row=master_rows + 3, column=self.psc_info.grid_info()['column'])
        self.hyp_user_manual.grid(
            row=master_rows + 3, column=self.hyp_user_manual.grid_info()['column'])
        self.version_tool_lbl.grid(
            row=master_rows + 3, column=self.version_tool_lbl.grid_info()['column'])

        return None

    def add_entry_gen_percent(self, row, col):
        """
			Function to add the text entry row for inserting fault times
		:param int row:  Row number to use
		:param int col:  Column number to use
		:return None:
		"""
        # Label for what is included in entry
        lbl = ttk.Label(
            master=self.gen_labelframe, text='% of Generator Maximum Output:', style=self.styles.label_general)
        lbl.grid(row=row, column=col, rowspan=1, sticky=Tk.W, padx=self.xpad, pady=self.ypad)
        # Set initial value for variable
        self.gen_percent_var.set(100.0)
        # Add entry box
        self.gen_percent_entry = Tk.Entry(master=self.gen_labelframe, textvariable=self.gen_percent_var)
        self.gen_percent_entry.grid(row=row + 1, column=col, sticky=Tk.W, padx=self.xpad, pady=self.ypad)
        self.gen_percent_entry.config(state=Tk.DISABLED)
        # CreateToolTip(widget=self.entry_fault_times, text=(
        # 	'Enter the durations after the fault the current should be calculated for.\n'
        # 	'Multiple values can be input in a list.'
        # ))
        return None

    def add_psc_logo_wm(self):
        """
			Function just adds the PSC logo to the windows manager in GUI
		:return: None
		"""
        # Create the PSC logo for including in the windows manager
        psc_logo_wm = Tk.PhotoImage(file=constants.GUI.img_pth_window)
        # noinspection PyProtectedMember
        self.master.tk.call('wm', 'iconphoto', self.master._w, psc_logo_wm)
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
        img = Image.open(constants.GUI.img_pth_main)
        img.thumbnail(constants.GUI.img_size)
        img = ImageTk.PhotoImage(img)
        psc_logo = Tk.Label(self.master, image=img, cursor='hand2', justify='center', compound='top')
        psc_logo.photo = img
        psc_logo.grid(row=row, column=col, columnspan=3, rowspan=3, pady=self.ypad)
        psc_logo.bind('<Button - 1>', lambda e: webbrowser.open_new('https://www.pscconsulting.com/'))
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
            self.master, text=constants.GUI.psc_phone, justify='left', font=constants.GUI.psc_font,
            foreground=constants.GUI.psc_color_grey, background=constants.GUIDefaultValues.color_main_window
        )
        self.psc_info.grid(row=row, column=col, columnspan=1)
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
        sep.grid(row=row, sticky=Tk.W + Tk.E, columnspan=col_span, pady=self.ypad)
        return None

    def select_sav_case(self):
        """
			Function to allow the user to select the SAV case to run
		:return: None
		"""
        # Ask user to select file(s) or folders based on <.bo_files>
        file_path = tkFileDialog.askopenfilename(
            initialdir=self.results_pth,
            filetypes=constants.General.sav_types,
            title='Select SAV case...'
        )

        # only do something if a file path was selected, if used selected cancel nothing will happen.
        if file_path:
            # Import busbar list from file assuming it is first column and append to existing list
            self.sav_case = file_path
            lbl_sav_button = 'SAV case = {}'.format(os.path.basename(self.sav_case))

            # Update command and radio button status
            self.cmd_select_sav_case.config(text=lbl_sav_button)
            enable_radio_buttons(self.load_radio_btn_list)
            enable_radio_buttons(self.gen_radio_btn_list)
            self.cmd_scale_load_gen.configure(state=Tk.NORMAL)

            # # If successful then get the zones and ratings from the loaded PSSE case and enable the buttons

            self.psse_con = psse.PsseControl()

            self.psse_con.load_data_case(pth_sav=self.sav_case)
            # psse_con.change_output(destination=False)

            self.zone_data = psse.ZoneData()
            self.zone_dict = self.zone_data.zone_dict

            self.psse_con.update_psse_gui()

        return None

    def import_load_estimates_xl(self):
        """
			Function to allow the user to select the SAV case to run
		:return: None
		"""
        # Ask user to select file(s) or folders based on <.bo_files>
        file_path = tkFileDialog.askopenfilename(
            initialdir=self.results_pth,
            filetypes=constants.General.file_types,
            title='Select SSE Load Estimates Spreadsheet'
        )
        # todo: @ DM: add the function to identify whether a dill exist and if it's newer than the file selected
        if file_path:
            # set load estimates to file path
            load_estimates_xl = file_path

            # process excel file
            # df_modified = dataframe_maker_modifier.main(load_estimates_xl,
            #                                             dill=True)  # if dill is true it would dill the dfs
            # todo:(done below) add one function to load a dilled df_modified
            df_modified = common_functions.load_dill(
                variable_name=[common_functions.folder_file_names.dill_modified_data],
                dill_folder_name=common_functions.folder_file_names.dill_folder)
            df_bad_data = common_functions.load_dill(
                variable_name=[common_functions.folder_file_names.dill_bad_data_name],
                dill_folder_name=common_functions.folder_file_names.dill_folder)

            if len(df_bad_data) > 0:
                constants.General.loads_complete = False
            else:
                constants.General.loads_complete = True

            constants.General.xl_file_name = common_functions.ExcelFileNames.bad_data_excel_name

            # TODO: To be completed
            self.set_gui_param(df_modified)

            # Update command and radio button status
            self.cmd_select_sav_case.configure(state=Tk.NORMAL)

            # Get just file name rather than full path
            xl_file_name = os.path.basename(load_estimates_xl)
            self.current_xl_path_lbl.configure(text=xl_file_name)

            if constants.General.loads_complete:
                self.load_complete_lbl_t_f.configure(text=constants.General.loads_complete_t_str, foreground='green')
                self.load_complete_lbl_t_f.unbind("<Button-1>")
            else:
                self.load_complete_lbl_t_f.configure(
                    text=constants.General.loads_complete_f_str, foreground='red', cursor='hand2')

                # file_path = os.path.abspath(os.path.join(
                #     load_est.PACKAGE_PATH, '..',
                #     constants.XlFileConstants.params_folder,
                #     constants.XlFileConstants.xl_checks_file_name))

                file_path = os.path.abspath(os.path.join(
                    load_est.PACKAGE_PATH,
                    common_functions.ExcelFileNames.data_comparison_excel_name))

                self.load_complete_lbl_t_f.bind("<Button-1>", lambda e: webbrowser.open_new(file_path))

            # update GUI variables
            # self.load_gsps = constants.General.scalable_GSP_list
            self.load_year_om.configure(values=self.load_year_list)
            # self.load_demand_scaling_om.configure(values=constants.General.demand_scaling_list)

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

    def set_gui_param(self, df):
        """
			Function will set extract the data from the DataFrame with regards to the Loads available from the
			load estimates
		:param pd.DataFrame df:  Raw dataframe that has been processed
		:return None:
		"""

        # Find the years being considered for the forecast
        self.load_year_list = common_functions.adjust_years(headers_list=list(df.columns))

        # The list of seasons are declared initially as they are constant

        # Retrieve unique list of GSPs from the DataFrame
        #   set() will reduce this to only a unique list
        #   list() will convert back to a list so that it can be used
        self.load_gsps = list(set(df.loc[:, common_functions.Headers.gsp]))
        k = 1

        # TODO: @DM - At this point may be necessary to actually force the GUI to update, need to
        # TODO: @DM... check the GUI combobox behaviour
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
        self.wait_time = 500  # milliseconds
        self.wrap_length = 450  # pixels
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.enter)
        self.widget.bind("<Leave>", self.leave)
        self.widget.bind("<ButtonPress>", self.leave)
        self.id = None
        self.tw = None

    # noinspection PyUnusedLocal
    def enter(self, event=None):
        del event
        self.schedule()

    # noinspection PyUnusedLocal
    def leave(self, event=None):
        del event
        self.unschedule()
        self.hidetip()

    # noinspection PyUnusedLocal
    def schedule(self, event=None):
        del event
        self.unschedule()
        self.id = self.widget.after(self.wait_time, self.showtip)

    # noinspection PyUnusedLocal
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

#!/usr/bin/env python
# -*- coding: utf-8 -*-
import tkinter as tk
import customtkinter as ctk
from PIL import Image, ImageTk
from tkinter.filedialog import asksaveasfilename
from datetime import datetime, timedelta
from utilities.regler import *
from utilities.data_functions import *
import time
import json
import os

save_timer = time.time()
write_header = 1

class TKH:
    def __init__(self, tfh_obj, json_name=False):
        # Store tfh_obj for later use
        self.tfh_obj = tfh_obj
        self.write_header = True  # Flag to control header writing
        self.save_timer = time.time()

        # Load configuration
        self.config = self.get_config(json_name)
        if not self.config:
            raise ValueError("Configuration could not be loaded")

        # Initialize window and other components
        self.window = self.initialize_window()
        self.set_background_image()
        self.labels = {}  # Initialize the labels dictionary
        self.entries = {}  # Initialize the entries dictionary
        self.buttons = {}  # Initialize the buttons dictionary
        self.controller = {}  # Initialize the controller dictionary

        # Create GUI components
        self.create_frames()
        self.create_entries(tfh_obj)
        self.create_labels(tfh_obj)
        self.create_buttons(tfh_obj)
        self.setup_controller(tfh_obj)
    
    def _create_label(self, parent, text, font_size, x=None, y=None, grid_opts=None, **kwargs):
        label = ctk.CTkLabel(parent, font=('Arial', font_size), text=text, **kwargs, bg_color='white')
        if grid_opts:
            label.grid(**grid_opts)
        else:
            label.place(x=x, y=y)
        return label
    
    def _create_button(self, parent, text, command, x=None, y=None, grid_opts=None, **kwargs):
        button = ctk.CTkButton(parent, text=text, command=command, **kwargs)
        if grid_opts:
            button.grid(**grid_opts)
        else:
            button.place(x=x, y=y)
        return button
    
    def _create_entry(self, parent, default_text, x=None, y=None, grid_opts=None, **kwargs):
        entry = ctk.CTkEntry(parent, **kwargs)
        entry.insert(0, str(default_text))
        if grid_opts:
            entry.grid(**grid_opts)
        else:
            entry.place(x=x, y=y)
        return entry

    def get_config(self, config_name):
            try:
                if config_name:
                    with open(f'./json_files/{config_name}.json', 'r') as config_file:
                        return json.load(config_file)
                else:
                    import config as cfg
                    return cfg.tkinter
            except (FileNotFoundError, json.JSONDecodeError) as e:
                print(f"Error loading config: {e}")
                return None


    def initialize_window(self):
            window = ctk.CTk()
            ctk.set_appearance_mode("light")
            
            # Use values from config or defaults
            scrW = self.config.get('TKINTER', {}).get('screen_width', 1280)
            scrH = self.config.get('TKINTER', {}).get('screen_height', 720)
            
            window.geometry(f"{scrW}x{scrH}")
            window.title(self.config['TKINTER'].get('Name', 'Default Title'))
            window.configure(bg=self.config['TKINTER'].get('background-color', '#FFFFFF'))
            #window.attributes('-fullscreen', self.config['TKINTER'].get('fullscreen', True))
            
            return window

    def set_background_image(self):
            bg_image_path = self.config['Background'].get('name', 'default_bg.png')
            bg_width = int(self.config['Background'].get('width', 1280))
            bg_height = int(self.config['Background'].get('height', 720))
            
            try:
                bg_image = ctk.CTkImage(Image.open(bg_image_path), size=(bg_width, bg_height))
                label_background = ctk.CTkLabel(self.window, image=bg_image, text="")
                label_background.place(
                    x=self.config['Background'].get('x', 0),
                    y=self.config['Background'].get('y', 0)
                )
                label_background.lower()
            except FileNotFoundError:
                print(f"Background image {bg_image_path} not found.")

    def create_frames(self):
        frames_dict = {}

        for frame_name, frame_config in self.config.get('Frames', {}).items():
            if frame_config.get('enabled', False):  # Erstelle den Frame nur, wenn "enabled" auf True gesetzt ist
                frames_dict[frame_name] = ctk.CTkFrame(
                    self.window,
                    fg_color=frame_config.get('fg_color', '#FFFFFF'),
                    border_color=frame_config.get('border_color', '#000000'),
                    border_width=frame_config.get('border_width', 5)
                )
                if 'x' in frame_config and 'y' in frame_config:
                    frames_dict[frame_name].place(
                        x=frame_config['x'],
                        y=frame_config['y']
                    )
                else:
                    frames_dict[frame_name].grid(
                        padx=frame_config.get('padx', 20),
                        pady=frame_config.get('pady', 20)
                    )
                
                # Erstelle ein Label im Frame, wenn ein Titel in der Konfiguration vorhanden ist
                if 'title' in frame_config:
                    name_frame = ctk.CTkLabel(
                        frames_dict[frame_name],
                        font=('Arial', 20),
                        text=frame_config['title']
                    )
                    name_frame.grid(
                        column=0, columnspan=2, row=0,
                        ipadx=7, ipady=7, pady=7, padx=7, sticky="E"
                    ) 
 
        # Speichern der Frames im Klassenattribut
        self.frames = frames_dict


    def create_labels(self, tfh_obj):
        
        labels_dict = {# Dictionaries to store labels for different device types
            'MFC': {},  'Tc': {},   'Pressure': {},     'Vorgabe': {},  'FlowMeter': {},    'ExtInput': {}
        }
      
        index_counters = { # Index counters for each device type
            'MFC': 0,   'Tc': 0,    'Pressure': 0,      'Vorgabe': 0,   'FlowMeter': 0,     'ExtInput': 0
        }

        for control_name, control_rule in tfh_obj.config.items():
            device_type = control_rule.get("type")

            if device_type == "mfc":
                idx = index_counters['MFC']
                labels_dict['MFC'][idx] = self._create_label(
                    parent=self.window,
                    # parent=self.frames['mfc'],
                    text='0 mV',
                    font_size=18,
                    #grid_opts={'column': 4, 'row': idx + 1, 'ipadx': 7, 'ipady': 7, 'padx': 20, 'pady': 20}
                    x=control_rule.get("x")+25,
                    y=control_rule.get("y")+45,
                )
                """  self._create_label(
                    parent=self.window,
                    text=control_name,
                    font_size=18,
                    x=control_rule.get("x")+80,
                    y=control_rule.get("y")
                    #grid_opts={'column': 1, 'row': idx + 1, 'ipadx': 5, 'ipady': 7, 'padx': 10, 'pady': 20}
                ) """
                self._create_label(
                    parent=self.window,
                    text=control_rule["DeviceInfo"].get("unit", 'mV'),
                    font_size=18,
                    x=control_rule.get("x")+50,
                    y=control_rule.get("y")
                    #grid_opts={'column': 3, 'row': idx + 1, 'ipadx': 1, 'ipady': 7, 'padx': 20, 'pady': 20}
                )
                index_counters['MFC'] += 1

            elif device_type == "thermocouple":
                idx = index_counters['Tc']
                labels_dict['Tc'][idx] = self._create_label(
                    parent=self.window,
                    text='0 °C',
                    font_size=18,
                    x=control_rule.get("x"),
                    y=control_rule.get("y")
                )
                index_counters['Tc'] += 1

            elif device_type == "pressure":
                idx = index_counters['Pressure']
                labels_dict['Pressure'][idx] = self._create_label(
                    parent=self.window,
                    text='0 bar',
                    font_size=18,
                    x=control_rule.get("x"),
                    y=control_rule.get("y")
                )
                index_counters['Pressure'] += 1

            elif device_type == "FlowMeter":
                idx = index_counters['FlowMeter']
                labels_dict['FlowMeter'][idx] = self._create_label(
                    parent=self.window,
                    text='0 kg/h',
                    font_size=18,
                    x=control_rule.get("x"),
                    y=control_rule.get("y")
                )
                index_counters['FlowMeter'] += 1

            elif device_type == "Vorgabe":
                idx = index_counters['Vorgabe']
                labels_dict['Vorgabe'][idx] = self._create_label(
                    parent=self.window,
                    text=control_rule["DeviceInfo"].get("unit", ''),
                    font_size=18,
                    x=control_rule.get("x") + 55,
                    y=control_rule.get("y")
                )
                index_counters['Vorgabe'] += 1

            elif device_type == "ExtInput":
                idx = index_counters['ExtInput']
                labels_dict['ExtInput'][idx] = self._create_label(
                    parent=self.window,
                    text='0 mA',
                    font_size=18,
                    x=control_rule.get("x"),
                    y=control_rule.get("y")
                )
                # Additional label for power
                labels_dict['ExtInput'][idx + 1] = self._create_label(
                    parent=self.window,
                    text='0 Watt',
                    font_size=18,
                    x=control_rule.get("x"),
                    y=control_rule.get("y") + 40,
                    bg_color='white'
                )
                index_counters['ExtInput'] += 2  # Increment by 2 since two labels are added

        # Save all created labels into the class attribute
        self.labels = labels_dict

    def create_buttons(self, tfh_obj):
        # Example: Create a dictionary to hold button references
        buttons_dict = {}

        # Ventile und Schalter aus TinkerForge Config
        for control_name, control_rule in tfh_obj.config.items():
            device_type = control_rule.get("type")
            if device_type == "valve":
                buttons_dict[control_name] = ctk.CTkSwitch(
                        self.window,
                        text = control_name,
                        font = ('Arial', 16),
                        bg_color=self.config['TKINTER'].get('background-color', '#FFFFFF')
                    )
                buttons_dict[control_name].place(x=control_rule.get('x'), y=control_rule.get('y'))

        # Set Button
        buttons_dict['Set'] = self._create_button(
            parent=self.frames['control'],
            text='Set Values',
            command=lambda: self.set_data(),  # Replace with the correct method
            grid_opts={'column': 0, 'row': 1, 'ipadx': 8, 'ipady': 8, 'padx': 20, 'pady': 10},  # Added padx for left margin, pady for vertical spacing
            fg_color='brown',
            text_color='white'
        )

        # Save Switch and Get FileName Button, falls in der Konfiguration aktiviert
        if self.config['TKINTER'].get('has_save_function', False):
            buttons_dict['Save'] = ctk.CTkSwitch(
                self.frames.get('control', self.window),
                text="Speichern",
                font=('Arial', 16)
            )
            buttons_dict['Save'].grid(column=2, row=1, ipadx=7, ipady=7, padx=20, pady=10)

            buttons_dict['GetFile'] = self._create_button(
                parent=self.frames.get('control', self.window),
                text='Data File',
                command=self.get_file,
                grid_opts={'column': 0, 'row': 2, 'ipadx': 8, 'ipady': 8, 'padx': 20, 'pady': 20},
                fg_color='brown',
                text_color='white'
            )

        # Close Button, falls in der Konfiguration aktiviert
        if self.config['TKINTER'].get('has_close_button', False):
            close_img = ctk.CTkImage(Image.open(self.config['Close']['name']), size=(80, 80))
            buttons_dict['Exit'] = ctk.CTkButton(
                master=self.window,
                text="",
                command=self.window.destroy,
                fg_color='transparent',
                bg_color= 'white',
                hover_color='#F2F2F2',
                image=close_img
            )
            buttons_dict['Exit'].place(x=self.config['Close']['x'], y=self.config['Close']['y'])


        # Save the created buttons into the class attribute
        self.buttons = buttons_dict

    def create_entries(self, tfh_obj):
        entries_dict = {
            'MFC': {},
            'Vorgabe': {}
        }

        i_MFC = 0
        i_V = 0

        for control_name, control_rule in tfh_obj.config.items():
            device_type = control_rule.get("type")

            if device_type == "mfc":
                entries_dict['MFC'][i_MFC] = self._create_entry(
                    parent=self.window,
                    default_text="0",
                    x=control_rule.get("x"),
                    y=control_rule.get("y"),
                    #grid_opts={'column': 2, 'row': i_MFC + 1, 'ipadx': 5, 'ipady': 7},
                    font=('Arial', 18),
                    width=40,
                    fg_color='light blue'
                )
                i_MFC += 1

            elif device_type == "Vorgabe":
                entries_dict['Vorgabe'][i_V] = self._create_entry(
                    parent=self.window,
                    default_text="0",
                    x=control_rule.get("x"),
                    y=control_rule.get("y"),
                    font=('Arial', 18),
                    width=40,
                    fg_color='light blue'
                )
                i_V += 1

        # Save the entries and a placeholder for the save file
        self.entries = entries_dict
        self.entries['SaveFile'] = "../Daten/test.dat"

    def setup_controller(self, tfh_obj):
        # Initialize dictionaries for controllers
        controllers_dict = {
            'easy_PI': {}
        }

        i_PI = 0

        for control_name, control_rule in tfh_obj.config.items():
            device_type = control_rule.get("type")
            
            if device_type == "easy_PI":
                out_device = control_rule.get("output_device")
                out_channel = control_rule.get("output_channel")
                P_val = control_rule["DeviceInfo"].get("P_Value")
                I_val = control_rule["DeviceInfo"].get("I_Value")

                # Create the easy_PI controller
                if "extern" in control_rule.get("input_device", "").lower():
                    controllers_dict['easy_PI'][i_PI] = easy_PI(out_device, out_channel, "extern", 0, I_val, P_val)
                else:
                    in_device = tfh_obj.config[control_rule.get("input_device")].get("input_device")
                    controllers_dict['easy_PI'][i_PI] = easy_PI(out_device, out_channel, tfh_obj.inputs[in_device], 0, I_val, P_val)

                # Create an entry widget for the controller
                controllers_dict['easy_PI'][i_PI].entry = ctk.CTkEntry(
                    self.window,
                    font=('Arial', 16),
                    width=50,
                    fg_color='light blue'
                )
                controllers_dict['easy_PI'][i_PI].entry.place(x=control_rule.get("x"), y=control_rule.get("y"))

                # Create a label for the controller output
                controllers_dict['easy_PI'][i_PI].label = ctk.CTkLabel(
                    self.window,
                    font=('Arial', 18),
                    text='0 %',
                    bg_color='white'
                )
                controllers_dict['easy_PI'][i_PI].label.place(x=control_rule.get("x"), y=control_rule.get("y") + 35)

                i_PI += 1

        # Save the controllers into the class attribute
        self.controller = controllers_dict

    def save_values(self):
        i_MFC = 0
        i_Tc = 0
        i_PI = 0
        i_V = 0

        # Write header if necessary
        if self.write_header:
            write_device_informations(self, self.tfh_obj)         
            with open(self.entries['SaveFile'], 'a') as f:
                header_line = '### Device Names \n Zeitpunkt \t'
                for control_name, control_rule in self.tfh_obj.config.items():
                    device_type = control_rule.get("type")
                    if device_type == "mfc":
                        header_line += f"{control_name}_Soll \t{control_name}_Ist \t"
                    else:
                        header_line += f"{control_name}\t"
                header_line += '\n'
                f.write(header_line)
            self.write_header = False

        # Write data
        with open(self.entries['SaveFile'], 'a') as f:
            data_line = f" {datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')}\t"
            for control_name, control_rule in self.tfh_obj.config.items():
                input_channel = control_rule.get("input_channel")
                input_device_uid = control_rule.get("input_device")
                output_channel = control_rule.get("output_channel")
                output_device_uid = control_rule.get("output_device")
                gradient = control_rule["DeviceInfo"].get("gradient")
                y_axis = control_rule["DeviceInfo"].get("y-axis")
                device_type = control_rule.get("type")

                # Handle different device types
                if device_type == "thermocouple":
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    data_line += f"{input_val}\t"

                elif device_type == "pressure":
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    data_line += f"{input_val}\t"

                elif device_type == "Vorgabe":
                    data_line += f"{self.entries['Vorgabe'][i_V].get()} \t"
                    i_V += 1

                elif device_type == "FlowMeter":
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    data_line += f"{input_val}\t"

                elif device_type == "ExtInput":
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    data_line += f"{input_val}\t"

                elif device_type == "mfc":
                    if self.tfh_obj.operation_mode != 1:  # Not in debug mode
                        input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    else:
                        input_val = 0.0
                    data_line += f"{self.entries['MFC'][i_MFC].get()} \t{input_val}\t"
                    i_MFC += 1

                elif device_type == "easy_PI":
                    converted_value = self.controller['easy_PI'][i_PI].out * 100
                    self.controller['easy_PI'][i_PI].label.configure(text=f"{round(converted_value, 2)} {control_rule['DeviceInfo'].get('unit')}")
                    data_line += f"{converted_value}\t"
                    i_PI += 1

            data_line += '\n'
            f.write(data_line)

    def set_data(self):
        i_MFC = 0
        i_PI = 0
        entries = self.entries
        controller = self.controller

        for control_name, control_rule in self.tfh_obj.config.items():
            device_type = control_rule.get("type")
            output_channel = control_rule.get("output_channel")
            output_device_uid = control_rule.get("output_device")
            gradient = control_rule["DeviceInfo"].get("gradient")
            y_axis = control_rule["DeviceInfo"].get("y-axis")

            # Handle MFC devices
            if device_type == "mfc":
                if entries['MFC'][i_MFC].get() != '':
                    value = float(entries['MFC'][i_MFC].get()) / gradient + y_axis
                    if self.tfh_obj.operation_mode != 1:  # Check if not in debug mode
                        self.tfh_obj.outputs[output_device_uid].values[output_channel] = value
                i_MFC += 1

            # Handle easy_PI devices
            if device_type == "easy_PI":
                if controller['easy_PI'][i_PI].entry.get() != '':
                    value = float(controller['easy_PI'][i_PI].entry.get())
                    if not controller['easy_PI'][i_PI].running:
                        controller['easy_PI'][i_PI].start(value)
                    else:
                        controller['easy_PI'][i_PI].set_soll(value)
                i_PI += 1

    def get_file(self):
        # Method to handle file selection
        file_path = asksaveasfilename(defaultextension=".dat", initialdir="./Daten/")
        if file_path:
            self.entries['SaveFile'] = file_path

            # Update a label or other widget to reflect the new file path
            parent_folder = os.path.basename(os.path.dirname(file_path))
            file_name = os.path.basename(file_path)
            short_text = f"{parent_folder}/{file_name}"
            if 'Save' in self.labels:
                self.labels['Save'].configure(text=short_text)

    def run(self):
        self.window.mainloop()

    def start_loop(self):
        i_MFC = 0
        i_Tc = 0
        i_PI = 0
        i_p = 0
        i_exI = 0
        i_FI = 0

        # Iterate over the configuration to update labels and data
        for control_name, control_rule in self.tfh_obj.config.items():
            input_channel = control_rule.get("input_channel")
            input_device_uid = control_rule.get("input_device")
            output_channel = control_rule.get("output_channel")
            output_device_uid = control_rule.get("output_device")
            gradient = control_rule["DeviceInfo"].get("gradient")
            y_axis = control_rule["DeviceInfo"].get("y-axis")
            unit = control_rule["DeviceInfo"].get("unit")
            device_type = control_rule.get("type")

            # Handle thermocouple devices
            if device_type == "thermocouple":
                input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                self.labels['Tc'][i_Tc].configure(text=f"{round(input_val, 2)} {unit}")
                i_Tc += 1

            # Handle pressure devices
            elif device_type == "pressure":
                if self.tfh_obj.operation_mode != 1:  # Not in debug mode
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    converted_value = ((input_val/1e6 - y_axis) * gradient)
                    self.labels['Pressure'][i_p].configure(text=f"{round(converted_value, 2)} {unit}")
                i_p += 1

            # Handle FlowMeter devices
            elif device_type == "FlowMeter":
                if self.tfh_obj.operation_mode != 1:  # Not in debug mode
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    converted_value = 0 + (100 - 0) / (20 - 4) * (input_val / 1e6 - 4)
                    if converted_value < 0:
                        converted_value = 0
                    self.labels['FlowMeter'][i_FI].configure(text=f"{round(converted_value, 2)} {unit}")
                i_FI += 1

            # Handle MFC devices
            elif device_type == "mfc":
                if self.tfh_obj.operation_mode != 1:  # Not in debug mode
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    converted_value = (input_val - y_axis) * gradient
                    self.labels['MFC'][i_MFC].configure(text=f"{round(converted_value, 2)} {unit}")
                i_MFC += 1

            # Handle easy_PI controllers
            elif device_type == "easy_PI":
                self.controller['easy_PI'][i_PI].regeln()
                converted_value = self.controller['easy_PI'][i_PI].out * 100
                if control_rule["DeviceInfo"].get('Power', False):
                    Power = control_rule["DeviceInfo"].get('Power')
                    self.controller['easy_PI'][i_PI].label.configure(text=f"{round(converted_value*Power, 0)} {unit}")
                else:
                    self.controller['easy_PI'][i_PI].label.configure(text=f"{round(converted_value, 2)} {unit}")

                if control_rule.get("output_type") == "analog":
                    converted_value = (4 + (20 - 4) / 100 * converted_value) * 1000
                self.tfh_obj.outputs[output_device_uid].values[output_channel] = converted_value
                i_PI += 1

            # Handle external inputs
            elif device_type == "ExtInput":
                input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                self.labels['ExtInput'][i_exI].configure(text=f"{round(input_val / 1e6, 2)} mA")
                i_exI += 1
                # Display power value if it's a heater
                if control_rule["DeviceInfo"].get('Power', False):
                    converted_value = (input_val - y_axis) * gradient
                    if converted_value < 0:
                        converted_value = 0
                    self.labels['ExtInput'][i_exI].configure(text=f"{round(converted_value, 2)} {unit}")
                i_exI += 1

            # Handle external inputs
            elif device_type == "valve":
                if self.buttons[control_name].get() == 1:
                    self.tfh_obj.outputs[output_device_uid].values[output_channel] = True
                else:
                    self.tfh_obj.outputs[output_device_uid].values[output_channel] = False

        # Save values periodically
        if self.buttons['Save'].get() == 1 and time.time() - self.save_timer > 1:
            self.save_values()
            self.save_timer = time.time()

        # Call this method again after 50 ms to continue the loop
        self.window.after(50, self.start_loop)

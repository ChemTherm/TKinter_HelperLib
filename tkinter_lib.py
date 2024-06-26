#!/usr/bin/env python
# -*- coding: utf-8 -*-
import tkinter as tk
import customtkinter as ctk
from PIL import Image, ImageTk
from tkinter.filedialog import asksaveasfilename
from datetime import datetime, timedelta
from utilities.regler import *
import time
import json

save_timer = time.time()


class TKH:
    def __init__(self, json_name):

        with open('./json_files/' + json_name + '.json', 'r') as config_file:
            config = json.load(config_file)


        window = ctk.CTk()
        ctk.set_appearance_mode("light")
        scrW = window.winfo_screenwidth()
        scrH = window.winfo_screenheight()
        scrW = 1280
        scrH = 720
        window.geometry(f"{scrW}x{scrH}")
        window.title(config['TKINTER']['Name'])
        window.configure(bg=config['TKINTER']['background-color'])
        window.attributes('-fullscreen', True)    

        bg_image = ctk.CTkImage(Image.open(config['Background']['name']), size=(int(config['Background']['width']), int(config['Background']['height'])))
        label_background = ctk.CTkLabel(window, image=bg_image, text="")
        label_background.place(x=config['Background']['x'], y=config['Background']['y'])
        label_background.lower()

        
        #----------- Frames ----------
        frames ={}
        frames['control'] = ctk.CTkFrame(window, fg_color = config['TKINTER']['background-color'], border_color = config['TKINTER']['border-color'], border_width=5)
        frames['control'].grid(column=0, row=1, padx=20, pady=20, ipadx = 20, ipady = 15)

        name_Frame = ctk.CTkLabel( frames['control'], font = ('Arial',20), text='Steuerung')
        name_Frame.grid(column=0, columnspan =2, row=0, ipadx=7, ipady=7, pady =7, padx = 7, sticky = "E")

        frames['timer'] = ctk.CTkFrame(window, fg_color = config['TKINTER']['background-color'], border_color = config['TKINTER']['border-color'], border_width=5)
        frames['timer'].grid(column=1, row=0, padx=20, pady=20, ipadx = 20, ipady = 15)
        
        frames['mfc']=ctk.CTkFrame(window, fg_color = config['TKINTER']['background-color'], border_color = config['TKINTER']['border-color'], border_width=5)
        frames['mfc'].grid(column=0, row=3, padx=20, pady=20, ipadx = 20, ipady = 15)
        tkinter_data = {window, frames, config}


def setdata(tk_obj, tfh_obj):
    i_MFC = 0; i_Tc = 0; i_PI = 0
    entries = tk_obj.entries
    controller = tk_obj.controller
    for control_name, control_rule in tfh_obj.config.items():
        device_type = control_rule.get("type"); output_channel = control_rule.get("output_channel"); output_device_uid = control_rule.get("output_device")
        gradient = control_rule.get("gradient"); y_axis = control_rule.get("y-axis"); unit = control_rule.get("unit")
    
        if device_type == "mfc":
            if entries['MFC'][i_MFC].get() != '':    
                value =  float(entries['MFC'][i_MFC].get())/gradient  + y_axis
                if tfh_obj.operation_mode == 1: #debug Mode
                    continue
                tfh_obj.outputs[output_device_uid].values[output_channel] = value
            i_MFC = i_MFC+1 

        if device_type == "easy_PI":
            if controller['easy_PI'][i_PI].entry.get() != '':
                value =  float(controller['easy_PI'][i_PI].entry.get())
                if controller['easy_PI'][i_PI].running == False:
                    controller['easy_PI'][i_PI].start(value)
                controller_value = controller['easy_PI'][i_PI].out
                tfh_obj.outputs[output_device_uid].values[output_channel] = controller_value
            i_PI = i_PI+1  


def setup_gui(json_name):
    with open('./json_files/' + json_name + '.json', 'r') as config_file:
        config = json.load(config_file)


    window = ctk.CTk()
    ctk.set_appearance_mode("light")
    scrW = window.winfo_screenwidth()
    scrH = window.winfo_screenheight()
    scrW = config['Background']['width']
    scrH = config['Background']['height']
    window.geometry(f"{scrW}x{scrH}")
    window.title(config['TKINTER']['Name'])
    window.configure(bg=config['TKINTER']['background-color'])
    #window.attributes('-fullscreen', True)    

    bg_image = ctk.CTkImage(Image.open(config['Background']['name']), size=(int(config['Background']['width']), int(config['Background']['height'])))
    close_img = ctk.CTkImage(Image.open(config['Close']['name']),size=(80, 80))
    label_background = ctk.CTkLabel(window, image=bg_image, text="")
    label_background.place(x=config['Background']['x'], y=config['Background']['y'])
    label_background.lower()

    
    Exit_Button = ctk.CTkButton(master=window,text="", command=window.destroy, fg_color= 'transparent',  hover_color='#F2F2F2', image= close_img)
    Exit_Button.place(x=config['Close']['x'], y=config['Close']['y'])
    #----------- Frames ----------
    frames ={}
    frames['control'] = ctk.CTkFrame(window, fg_color = config['TKINTER']['background-color'], border_color = config['TKINTER']['border-color'], border_width=5)
    frames['control'].grid(column=0, row=1, padx=20, pady=20, ipadx = 20, ipady = 15)

    name_Frame = ctk.CTkLabel( frames['control'], font = ('Arial',20), text='Steuerung')
    name_Frame.grid(column=0, columnspan =2, row=0, ipadx=7, ipady=7, pady =7, padx = 7, sticky = "E")

    frames['timer'] = ctk.CTkFrame(window, fg_color = config['TKINTER']['background-color'], border_color = config['TKINTER']['border-color'], border_width=5)
    frames['timer'].grid(column=1, row=0, padx=20, pady=20, ipadx = 20, ipady = 15)
    
    frames['mfc']=ctk.CTkFrame(window, fg_color = config['TKINTER']['background-color'], border_color = config['TKINTER']['border-color'], border_width=5)
    frames['mfc'].grid(column=0, row=3, padx=20, pady=20, ipadx = 20, ipady = 15)
    tkinter = window
    tkinter.frames = frames
    tkinter.config = config

    return tkinter

def tk_loopNew(tk_obj, tfh_obj)   :
    global save_timer
    window = tk_obj
    labels = tk_obj.labels
    entries = tk_obj.entries
    controller = tk_obj.controller
    i_MFC = 0; i_Tc = 0;  i_PI = 0

    for control_name, control_rule in tfh_obj.config.items():
        input_channel = control_rule.get("input_channel")
        input_device_uid = control_rule.get("input_device")
        output_channel = control_rule.get("output_channel")
        output_device_uid = control_rule.get("output_device")
        gradient = control_rule.get("gradient")
        y_axis = control_rule.get("y-axis")
        unit = control_rule.get("unit")
        device_type = control_rule.get("type")        


        if device_type == "thermocouple":
            # Handle Output  
            input_val = tfh_obj.inputs[input_device_uid].values[input_channel]  
            labels['Tc'][i_Tc].configure(text=f"{round(input_val, 2)} " + unit )
            i_Tc = i_Tc + 1


        if device_type == "mfc":
           
            if tfh_obj.operation_mode == 1: #debug Mode
                input_val = 0.0      
            else:
                input_val = tfh_obj.inputs[input_device_uid].values[input_channel]
            converted_value = (input_val - y_axis) * gradient   
            labels['MFC'][i_MFC].configure(text=f"{round(converted_value, 2)} " + unit )
            i_MFC = i_MFC+1
            
        if device_type == "easy_PI":
            controller['easy_PI'][i_PI].regeln()
            converted_value = controller['easy_PI'][i_PI].out*100
            controller['easy_PI'][i_PI].label.configure(text=f"{round(converted_value, 2)} " + unit )
            i_PI = i_PI+1
    tk_obj.after(50, tk_loopNew, tk_obj, tfh_obj) 


def create_entries(tk_obj, tfh_obj):
    frames = tk_obj.frames
    tk_obj.entries = {}
    MFCs = {}
    i = 0
    for control_name, control_rule in tfh_obj.config.items():
        device_type = control_rule.get("type")
        if device_type == "mfc":
            MFCs[i] = tk.Entry(frames['mfc'], font=('Arial', 16), width=6, bg='light blue')
            MFCs[i].grid(column=2, row=i+1, ipadx=5, ipady=7)
            i = i+1

    
    tk_obj.entries = {'MFC' : MFCs}
    return tk_obj

def create_labels(tk_obj, tfh_obj):
    frames = tk_obj.frames
    window = tk_obj
    tk_obj.labels = {}
    MFCs = {}; Tcs = {}
    i_MFC = 0; i_Tc = 0
    name_Frame = ctk.CTkLabel( frames['mfc'], font = ('Arial',20), text='MFC Steuerung')
    name_Frame.grid(column=0, columnspan =3, row=0, ipadx=7, ipady=7, pady =7, padx = 7, sticky = "E")
    
    name_MFC={};  unit_MFC={}; 

    for control_name, control_rule in tfh_obj.config.items():
        device_type = control_rule.get("type")
        if device_type == "mfc":
            gradient = control_rule.get("gradient"); unit = control_rule.get("unit")
            name_MFC[i_MFC]= ctk.CTkLabel( frames['mfc'], font = ('Arial',16), text=control_name)
            name_MFC[i_MFC].grid(column=1, row=i_MFC+1, ipadx=5, ipady=7)
            
            unit_MFC[i_MFC]= ctk.CTkLabel( frames['mfc'], font = ('Arial',16), text=' mV')
            unit_MFC[i_MFC].grid(column=3, row=i_MFC+1, ipadx=1, ipady=7)
            if gradient > 0: # Set MFC with unit
                unit_MFC[i_MFC].configure(text= unit)

            MFCs[i_MFC] = ctk.CTkLabel(frames['mfc'], font = ('Arial',16), text='0 mV')
            MFCs[i_MFC].grid(column=4, row=i_MFC+1, ipadx=7, ipady=7)
            i_MFC = i_MFC+1 # inkrement 
        elif device_type == "thermocouple":
            Tcs[i_Tc] = ctk.CTkLabel(window, font = ('Arial',16), text='0 °C')
            Tcs[i_Tc].place(x=  control_rule.get("x"), y= control_rule.get("y"))
            i_Tc = i_Tc+1 # inkrement 
        
    tk_obj.labels = {'MFC' : MFCs, 'Tc': Tcs}

    return tk_obj

def setup_controller(tk_obj, tfh_obj): #Function for Tkinter
    PIs = {};     i_PI = 0

    for control_name, control_rule in tfh_obj.config.items():
        device_type = control_rule.get("type")
        if device_type == "easy_PI":
            out_device = control_rule.get("output_device")
            out_channel = control_rule.get("output_channel")
            in_device = tfh_obj.config[control_rule.get("input_device")].get("input_device")

            P_val = control_rule.get("P_Value")
            I_val = control_rule.get("I_Value")
            PIs[i_PI] = easy_PI(out_device, out_channel, tfh_obj.inputs[in_device], 0, P_val, I_val)
            PIs[i_PI].entry = tk.Entry(tk_obj, font=('Arial', 16), width=6, bg='light blue')
            PIs[i_PI].entry.place(x=  control_rule.get("x"), y= control_rule.get("y"))
            PIs[i_PI].label =  ctk.CTkLabel(tk_obj, font = ('Arial',16), text='0 %')
            PIs[i_PI].label.place(x=  control_rule.get("x"), y= control_rule.get("y") + 35)
            
            i_PI = i_PI+1 # inkrement 


            
    tk_obj.controller = {'easy_PI' : PIs}

    return tk_obj        


def create_buttons(tk_obj, tfh_obj):
 #----------- Buttons -----------
    set_button = tk.Button(tk_obj.frames['control'],text='Set Values', command = lambda: setdata(tk_obj, tfh_obj), bg='brown', fg='white')
    set_button.grid(column=1, row=0, ipadx=8, ipady=8)

    tk_obj.buttons = {'Set': set_button}
    return tk_obj

def getfile(entry_list, label_list):
    entry_list['File'] = asksaveasfilename(defaultextension = ".dat", initialdir= "D:/Daten/")
    label_list['Save'].configure(text=entry_list['File'])

def setup_frames_labels_buttons(window, frames, img, device_list, entry_list, label_list):
    
    save_switch =  ctk.CTkSwitch(frames['control'], font=('Arial', 16), text="Speichern")
    save_switch.grid(column=2, row=2, ipadx=7, ipady=7)
    
    label_list['Save'] = ctk.CTkLabel(frames['control'], font = ('Arial',16), text=entry_list['File'])
    label_list['Save'].grid(column=0, columnspan = 4, row=3, ipadx=7, ipady=7)
    
    get_filename = ctk.CTkButton(frames['control'], text = 'Data File', command = lambda: getfile(entry_list, label_list), fg_color = 'brown')
    get_filename.grid(column= 3, row = 2, ipadx=7, ipady=7)

    set_button = ctk.CTkButton(frames['control'],text='Set Values', command = lambda: setdata(device_list, entry_list), fg_color = 'brown')
    set_button.grid(column=0 ,columnspan = 3, row=1, ipadx=8, ipady=8, padx = 7, pady = 7) 

    close_img = ctk.CTkImage(Image.open(img['Close']['name']), size=(80, 80))
    exit_button = ctk.CTkButton(master = window, text="", command=window.destroy, fg_color='transparent', hover_color='#F2F2F2', image=close_img)
    exit_button.place(x=img['Close']['x'], y=img['Close']['y'])
    
    return  save_switch 

def save_values(device_list, label_list, entry_list):
    tc_list = device_list['T']
    hp_list = device_list['HP']
    mfc_list = device_list['MFC']
    ABB_list = device_list['ABB']
    pressure_list = device_list['P']

    with open(entry_list['File'], 'a') as f:
        line = ' ' + datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f ")+ '\t'
        for index, tc_instance in enumerate(tc_list):
            line += str(tc_instance.t) + '\t'
        for index, hp_instance in enumerate(hp_list):
            line += str(hp_instance.pwroutput) + '\t'
        for index, mfc_instance in enumerate(mfc_list):
            if mfc_instance.m > 0:
                line += str(entry_list['MFC'][index].get()) + ' \t '+ str(mfc_instance.value) + ' \t '
            else:
                line += str(entry_list['MFC'][index].get()) + ' \t '+ str(mfc_instance.Voltage) + ' \t '
        for index, pressure_instance in enumerate(pressure_list):
            if pressure_instance.m > 0:
                line += str(pressure_instance.current) + '\t'
            else:
                line += str(pressure_instance.value) + '\t'
        line += ' \n'
        f.writelines(line)



def tk_loop(window, device_list, label_list, entry_list) :
    global save_timer
    
    label_T_ist = label_list['T']
    tc_list = device_list['T']
    label_HP_ist = label_list['HP']
    hp_list = device_list['HP']
    label_MFC = label_list['MFC']
    mfc_list = device_list['MFC']
    save_switch = entry_list['Save']
    pressure_list = device_list['P']
    label_pressure = label_list['P']


    for index, tc_instance in enumerate(tc_list):
        label_T_ist[index].configure(text=f"{round(tc_instance.t, 2)} °C")

    for index, hp_instance in enumerate(hp_list):
        label_HP_ist[index].configure(text=f"{round(hp_instance.pwroutput * 100, 2)} %")
        hp_instance.regeln()


    for index, pressure_instance in enumerate(pressure_list):
        pressure_instance.get()
        label_pressure[index].configure(text=f"{round(pressure_instance.current, 2)} mV")
        if pressure_instance.m > 0:
            label_pressure[index].configure(text=f"{round(pressure_instance.value, 2)} " + pressure_instance.unit )

    for index, mfc_instance in enumerate(mfc_list):
        mfc_instance.get()
        label_MFC[index].configure(text=f"{round(mfc_instance.Voltage, 2)} mV")
        if mfc_instance.m > 0:
            label_MFC[index].configure(text=f"{round(mfc_instance.value, 2)} " + mfc_instance.unit )

    if save_switch.get() == 1 and save_timer - time.time()< 0:
        save_values(device_list, label_list, entry_list)
        save_timer = time.time() + 1000/1000
    


    window.after(50, tk_loop, window, device_list, label_list, entry_list) 


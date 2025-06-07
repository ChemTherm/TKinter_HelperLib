#!/usr/bin/env python
# -*- coding: utf-8 -*-

import tkinter as tk
import customtkinter as ctk
from PIL import Image, ImageTk
from tkinter.filedialog import asksaveasfilename, askopenfilename
from tkinter import messagebox
from datetime import datetime, timedelta
from utilities.regler import *
from utilities.data_functions import *

from utilities.regler import DirectHeatController
import time
import json
import os
import openpyxl

# Globaler Timer für Excel-Logging
save_timer = time.time()
write_header = 1

def Excel_timing(sheet, section, t0):
    """
    Liest aus der gegebenen Zeile (section) des Excel-Sheets:
      - Die erste Zelle enthält die Zeitdauer (in Sekunden) des Abschnitts.
      - Die restlichen Zellen enthalten entweder einen direkten Sollwert (z.B. 200) oder
        einen Bereich im Format "Start-End" (z.B. "100-200"), für den ein interpolierter Sollwert ermittelt wird.
    
    Parameter:
      sheet   : Das geöffnete Excel-Arbeitsblatt (openpyxl Worksheet)
      section : Die Zeilennummer, die aktuell abgearbeitet wird
      t0      : Der Zeitpunkt, an dem der aktuelle Abschnitt begonnen hat (time.time())
    
    Rückgabe:
      output    : Dictionary, z. B. { 'Heater_1': <aktueller Sollwert>, ... }
      section   : ggf. aktualisierte Zeilennummer (nächster Abschnitt, falls abgelaufen)
      t_section : Verbleibende Zeit des aktuellen Abschnitts (Sekunden)
      t0        : ggf. aktualisierter Startzeitpunkt für den neuen Abschnitt
    """

    # Lese alle Zellen der aktuellen Zeile (ohne Filter, damit die Spaltenreihenfolge erhalten bleibt)
    row = sheet[section]
    values = [cell.value for cell in row]

    # Die erste Zelle enthält die Zeitdauer des Abschnitts
    try:
        section_time = float(str(values[0]).replace(',', '.'))
    except (ValueError, TypeError):
        raise ValueError(f"Ungültiger Zeitwert in Zeile {section}: {values[0]}")
    
    # Berechne die verstrichene Zeit seit Beginn des Abschnitts und die verbleibende Zeit
    elapsed = time.time() - t0
    t_section = section_time - elapsed

    # Hole die Header aus Zeile 1 (angenommen, hier stehen die Spaltenüberschriften)
    header = [cell.value for cell in sheet[2]]

    output = {}
    # Gehe alle Zellen ab der 2. Spalte durch, da die 1. Spalte die Zeitdauer ist.
    for i, val in enumerate(values[1:], start=1):
        # Hole den entsprechenden Spaltennamen, falls vorhanden
        key = header[i] if i < len(header) else f"Column_{i}"
        
        # Falls der Zellinhalt ein String mit '-' ist, wird ein Bereich erwartet
        if isinstance(val, str) and '-' in val:
            parts = val.split('-')
            try:
                lower = float(parts[0].replace(',', '.').strip())
                upper = float(parts[1].replace(',', '.').strip())
                #print(f"Interpolating {lower} to {upper} at {elapsed} / {section_time}")
            except ValueError:
                # Bei einem Umrechnungsproblem wird der Originalinhalt übernommen
                output[key] = val
                continue
            # Berechne den Fortschritt im aktuellen Abschnitt, abgeklammert zwischen 0 und 1
            progress = min(max(elapsed / section_time, 0), 1)
            interp_val = lower + (upper - lower) * progress
            output[key] = interp_val
            #print(f"Interpolated value: {interp_val}")
            #print(key, lower, upper, progress)
        else:
            # Falls es sich um eine einzelne Zahl handelt, versuche sie in float zu konvertieren
            try:
                output[key] = float(str(val).replace(',', '.'))
            except (ValueError, TypeError):
                output[key] = val

    # Wenn die Zeit des aktuellen Abschnitts abgelaufen ist, gehe zum nächsten Abschnitt und setze t0 zurück.
    if t_section < 0:
        section += 1
        t0 = time.time()

    return output, section, t_section, t0


class TKH:
    
    """
    Diese Klasse stellt die GUI zur Verfügung und verbindet Tkinter/CustomTkinter mit
    den Konfigurationsdaten, Messwerten und Steuerungsfunktionen.
    
    Die Klasse übernimmt u.a.:
      - Laden und Parsen der Konfiguration (JSON oder config-Modul)
      - Erstellen von Fenstern, Frames, Labels, Buttons und Eingabefeldern
      - Einfügen von Hintergrundbildern und weiteren Grafiken
      - Regelmäßiges Aktualisieren und Speichern der Messwerte
    """
    def __init__(self, tfh_obj, modbus_obj, json_name=False):
        # Objekte für Daten/Steuerung speichern
        self.tfh_obj = tfh_obj
        self.modbus_obj = modbus_obj
        self.write_header = True
        self.save_timer = time.time()
        self.running_excel = 0
        
        # Konfiguration laden (JSON oder über ein config-Modul)
        self.config = self.get_config(json_name)
        if not self.config:
            raise ValueError("Configuration could not be loaded")
        
        # Fenster und GUI-Komponenten initialisieren
        self.window = self.initialize_window()
        self.set_all_pictures()
        
        # Dictionaries zum Speichern von Widgets
        self.labels = {}
        self.entries = {}
        self.buttons = {}
        self.controller = {}
        
        # Frames, Eingabefelder, Labels, Buttons und Controller erstellen
        self.create_frames()
        self.create_entries(tfh_obj)
        self.create_labels(tfh_obj)
        self.create_buttons(tfh_obj)
        self.setup_controller(tfh_obj)
    
    # --- Hilfsfunktionen zum Erzeugen von Widgets ---
    def _create_label(self, parent, text, font_size, x=None, y=None, grid_opts=None, **kwargs):
        """
        Erzeugt ein Label mit dem angegebenen Parent, Text und Schriftgröße.
        Platzierung erfolgt entweder über .grid() oder .place().
        """
        label = ctk.CTkLabel(parent, font=('Arial', font_size), text=text, bg_color='white', **kwargs)
        if grid_opts:
            label.grid(**grid_opts)
        else:
            label.place(x=x, y=y)
        return label
    
    def _create_button(self, parent, text, command, x=None, y=None, grid_opts=None, **kwargs):
        """
        Erzeugt einen Button mit dem angegebenen Parent, Text und Callback.
        """
        button = ctk.CTkButton(parent, text=text, command=command, **kwargs)
        if grid_opts:
            button.grid(**grid_opts)
        else:
            button.place(x=x, y=y)
        return button
    
    def _create_entry(self, parent, default_text, x=None, y=None, grid_opts=None, **kwargs):
        """
        Erzeugt ein Eingabefeld (Entry), füllt es mit dem Default-Text und platziert es.
        """
        entry = ctk.CTkEntry(parent, **kwargs)
        entry.insert(0, str(default_text))
        if grid_opts:
            entry.grid(**grid_opts)
        else:
            entry.place(x=x, y=y)
        return entry
    
    # --- Konfiguration laden und Fenster initialisieren ---
    def get_config(self, config_name):
        """
        Lädt die Konfiguration entweder aus einer JSON-Datei oder aus dem config-Modul.
        
        :param config_name: Name der JSON-Datei (ohne Endung) oder False, um das config-Modul zu verwenden.
        :return: Konfigurationsdictionary oder None bei Fehler.
        """
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
        """
        Initialisiert das Hauptfenster basierend auf Konfigurationsparametern.
        """
        window = ctk.CTk()
        ctk.set_appearance_mode("light")
        
        # Bildschirmgröße aus der Konfiguration oder Standardwerte
        scrW = self.config.get('TKINTER', {}).get('screen_width', 1280)
        scrH = self.config.get('TKINTER', {}).get('screen_height', 720)
        
        window.geometry(f"{scrW}x{scrH}")
        window.title(self.config['TKINTER'].get('Name', 'Default Title'))
        window.configure(bg=self.config['TKINTER'].get('background-color', '#FFFFFF'))
        window.attributes('-fullscreen', self.config['TKINTER'].get('fullscreen', True))
        
        return window

    # --- Bilder in das Fenster einfügen ---
    def set_all_pictures(self):
        """
        Durchläuft alle Konfigurationseinträge und fügt jene mit dem Typ "picture" als Bild ein.
        
        Unterstützt werden Einträge, die entweder den Schlüssel "name" oder "png" für den Bildpfad enthalten.
        Größe und Position werden aus der Konfiguration ausgelesen.
        """
        for key, pic_conf in self.config.items():
            if pic_conf.get("type") == "picture":
                # Ermittele Bildpfad (name oder png)
                image_path = pic_conf.get("name") or pic_conf.get("png")
                if not image_path:
                    print(f"Kein Bildpfad in der Konfiguration für {key} gefunden.")
                    continue

                width = int(pic_conf.get("width", 0))
                height = int(pic_conf.get("height", 0))
                x = int(pic_conf.get("x", 0))
                y = int(pic_conf.get("y", 0))
                
                try:
                    image = ctk.CTkImage(Image.open(image_path), size=(width, height))
                    label = ctk.CTkLabel(self.window, image=image, text="")
                    label.place(x=x, y=y)
                    label.lower()  # Hintergrundbild nach hinten verschieben
                    
                    # Referenz speichern, um Garbage Collection zu verhindern
                    if not hasattr(self, "_image_refs"):
                        self._image_refs = []
                    self._image_refs.append(image)
                except FileNotFoundError:
                    print(f"Bilddatei '{image_path}' für {key} nicht gefunden.")

    # --- Frames erstellen ---
    def create_frames(self):
        """
        Erstellt Frames gemäß der Konfiguration.
        
        Falls im Frame eine Platzierung (x, y) angegeben ist, wird .place() verwendet,
        ansonsten .grid() mit Padding.
        Zusätzlich wird ein Titel-Label im Frame erzeugt, falls in der Konfiguration angegeben.
        """
        frames_dict = {}
        for frame_name, frame_config in self.config.get('Frames', {}).items():
            if frame_config.get('enabled', False):
                frames_dict[frame_name] = ctk.CTkFrame(
                    self.window,
                    fg_color=frame_config.get('fg_color', '#FFFFFF'),
                    border_color=frame_config.get('border_color', '#000000'),
                    border_width=frame_config.get('border_width', 5)
                )
                if 'x' in frame_config and 'y' in frame_config:
                    frames_dict[frame_name].place(x=frame_config['x'], y=frame_config['y'])
                else:
                    frames_dict[frame_name].grid(
                        padx=frame_config.get('padx', 20),
                        pady=frame_config.get('pady', 20)
                    )
                
                if 'title' in frame_config:
                    name_frame = ctk.CTkLabel(
                        frames_dict[frame_name],
                        font=('Arial', 20),
                        text=frame_config['title']
                    )
                    name_frame.grid(column=0, columnspan=2, row=0,
                                    ipadx=7, ipady=7, pady=7, padx=7, sticky="E")
        
        self.frames = frames_dict

    # --- Labels erstellen ---
    def create_labels(self, tfh_obj):
        """
        Erstellt und platziert Labels für unterschiedliche Gerätetypen.
        
        Verwendet werden zwei Konfigurationsquellen:
          - self.modbus_obj.config: Für Geräte, die über Modbus gesteuert werden.
          - tfh_obj.config: Für zusätzliche externe Geräte.
          
        Die Labels werden in einem Dictionary (self.labels) gespeichert.
        """
        # Initialisiere Dictionaries und Indexzähler für verschiedene Gerätetypen.
        labels_dict = {
            'mfc': {},
            'Tc': {},
            'Pressure': {},
            'Vorgabe': {},
            'FlowMeter': {},
            'ExtInput': {},
            'ExtOutput': {},
            'Modbus_Pump': {},
            'analytic': {}
        }
        index_counters = {key: 0 for key in labels_dict}

        # --- Verarbeitung der Modbus-Konfiguration ---
        for control_name, control_rule in self.modbus_obj.config.items():
            device_type = control_rule.get("type")
            if device_type == "mfc":
                idx = index_counters['mfc']
                # Ersetze Unterstriche im Namen durch Leerzeichen für eine bessere Anzeige
                display_text = control_name.replace("_", " ")
                if control_rule.get("Box") == 1:
                    parent_var = self.frames['mfc']
                    options = {'grid_opts': {'column': 1, 'row': idx + 1, 'ipadx': 7, 'ipady': 7, 'padx': 5, 'pady': 5}}
                else:
                    parent_var = self.window
                    options = {'x': control_rule.get("x") + 25, 'y': control_rule.get("y") + 45}
                
                # Erstes Label: Anzeigen des (formatierten) control_name
                labels_dict['mfc'][idx] = self._create_label(
                    parent=parent_var,
                    text=display_text,
                    font_size=18,
                    **options
                )
                
                # Zweites Label: Anzeige des Werts "0" plus Einheit aus DeviceInfo
                if control_rule.get("Box") == 1:
                    parent_var = self.frames['mfc']
                    options = {'grid_opts': {'column': 4, 'row': idx + 1, 'ipadx': 7, 'ipady': 7, 'padx': 20}}
                else:
                    parent_var = self.window
                    options = {'x': control_rule.get("x") + 25, 'y': control_rule.get("y") + 45}
                labels_dict['mfc'][idx] = self._create_label(
                    parent=parent_var,
                    text="0 " + control_rule["DeviceInfo"].get("unit"),
                    font_size=18,
                    **options
                )
                
                # Drittes Label: Anzeige der Einheit (Standard 'mV' falls nicht vorhanden)
                if control_rule.get("Box") == 1:
                    parent_var = self.frames['mfc']
                    options = {'grid_opts': {'column': 3, 'row': idx + 1, 'ipadx': 1, 'ipady': 7, 'padx': 20}}
                else:
                    parent_var = self.window
                    options = {'x': control_rule.get("x") + 50, 'y': control_rule.get("y")}
                self._create_label(
                    parent=parent_var,
                    text=control_rule["DeviceInfo"].get("unit", 'mV'),
                    font_size=18,
                    **options
                )
                index_counters['mfc'] += 1

        # --- Verarbeitung der tfh-Konfiguration (externe Geräte) ---
        for control_name, control_rule in tfh_obj.config.items():
            device_type = control_rule.get("type")
            if device_type == "mfc" and "modbus" not in control_rule.get("input_device", "").lower():
                idx = index_counters['mfc']
                # Erstes Label: Standard '0 mV'
                labels_dict['mfc'][idx] = self._create_label(
                    parent=self.window,
                    text='0 mV',
                    font_size=18,
                    x=control_rule.get("x") + 25,
                    y=control_rule.get("y") + 45,
                )
                # Zweites Label: Anzeige der Einheit
                self._create_label(
                    parent=self.window,
                    text=control_rule["DeviceInfo"].get("unit", 'mV'),
                    font_size=18,
                    x=control_rule.get("x") + 50,
                    y=control_rule.get("y")
                )
                index_counters['mfc'] += 1

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
            
            elif device_type == "analytic":
                idx = index_counters['analytic']
                labels_dict['analytic'][idx] = self._create_label(
                    parent=self.window,
                    text='0',
                    font_size=18,
                    x=control_rule.get("x"),
                    y=control_rule.get("y")
                )
                index_counters['analytic'] += 1

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

            elif device_type == "Modbus_Pump":
                idx = index_counters['Modbus_Pump']
                labels_dict['Modbus_Pump'][idx] = self._create_label(
                    parent=self.window,
                    text=control_rule["DeviceInfo"].get("unit", ''),
                    font_size=18,
                    x=control_rule.get("x") + 55,
                    y=control_rule.get("y")
                )
                index_counters['Modbus_Pump'] += 1

            elif device_type == "ExtInput":
                idx = index_counters['ExtInput']
                labels_dict['ExtInput'][idx] = self._create_label(
                    parent=self.window,
                    text='0 mA',
                    font_size=18,
                    x=control_rule.get("x"),
                    y=control_rule.get("y")
                )
                # Zusätzliches Label für Leistung (Watt)
                labels_dict['ExtInput'][idx + 1] = self._create_label(
                    parent=self.window,
                    text='0 Watt',
                    font_size=18,
                    x=control_rule.get("x"),
                    y=control_rule.get("y") + 40,
                    bg_color='white'
                )
                index_counters['ExtInput'] += 2

        # Timer-Label (nur wenn Excel-Funktion aktiviert)
        if self.config['TKINTER'].get('has_excel_function', False):
            labels_dict['Timer'] = self._create_label(
                parent=self.frames['control'],
                text='0 min',
                font_size=18,
                grid_opts={'column': 2, 'row': 0, 'ipadx': 2, 'ipady': 2, 'padx': 10, 'pady': 10},
            )
        
        self.labels = labels_dict

    # --- Buttons erstellen ---
    def create_buttons(self, tfh_obj):
        """
        Erzeugt Schalter und Buttons basierend auf der Konfiguration.
        
        Neben den Ventilen (Switches) werden auch Buttons für das Setzen von Werten,
        Excel-Funktionen, Dateiauswahl und das Schließen des Programms erstellt.
        """
        buttons_dict = {}

        # Ventile: Ersetze "_" im Namen durch Leerzeichen
        for control_name, control_rule in tfh_obj.config.items():
            if control_rule.get("type") == "valve":
                display_text = control_name.replace("_", " ")
                buttons_dict[control_name] = ctk.CTkSwitch(
                    self.window,
                    text=display_text,
                    font=('Arial', 16),
                    bg_color=self.config['TKINTER'].get('background-color', '#FFFFFF')
                )
                buttons_dict[control_name].place(x=control_rule.get('x'), y=control_rule.get('y'))
                

        # "Set Values"-Button
        buttons_dict['Set'] = self._create_button(
            parent=self.frames['control'],
            text='Set Values',
            command=lambda: self.set_data(),
            grid_opts={'column': 0, 'row': 1, 'ipadx': 8, 'ipady': 6, 'padx': 20, 'pady': 10},
            fg_color='brown',
            text_color='white'
        )
        
        # Excel-Buttons (falls aktiviert)
        if self.config['TKINTER'].get('has_excel_function', False):
            buttons_dict['StartExcel'] = self._create_button(
                parent=self.frames['control'],
                text='Start Excel',
                command=lambda: self.start_excel(),
                grid_opts={'column': 2, 'row': 3, 'ipadx': 8, 'ipady': 6, 'padx': 20, 'pady': 10},
                fg_color='brown',
                text_color='white'
            )
            buttons_dict['StopExcel'] = self._create_button(
                parent=self.frames['control'],
                text='Stop Excel',
                command=lambda: self.stop_excel(),
                grid_opts={'column': 2, 'row': 4, 'ipadx': 8, 'ipady': 6, 'padx': 20, 'pady': 10},
                fg_color='brown',
                text_color='white'
            )
            buttons_dict['GetExcel'] = self._create_button(
                parent=self.frames.get('control', self.window),
                text='Excel File',
                command=self.get_Excelfile,
                grid_opts={'column': 0, 'row': 3, 'ipadx': 8, 'ipady': 6, 'padx': 20, 'pady': 10},
                fg_color='brown',
                text_color='white'
            )
        
        # Speichern und Dateiauswahl (falls aktiviert)
        if self.config['TKINTER'].get('has_save_function', False):
            buttons_dict['Save'] = ctk.CTkSwitch(
                self.frames.get('control', self.window),
                text="Speichern",
                font=('Arial', 16)
            )
            buttons_dict['Save'].grid(column=2, row=2, ipadx=7, ipady=7, padx=20, pady=10)

            buttons_dict['GetFile'] = self._create_button(
                parent=self.frames.get('control', self.window),
                text='Data File',
                command=self.get_file,
                grid_opts={'column': 0, 'row': 2, 'ipadx': 8, 'ipady': 6, 'padx': 20, 'pady': 10},
                fg_color='brown',
                text_color='white'
            )

        # Schließen-Button (falls aktiviert)
        if self.config['TKINTER'].get('has_close_button', False):
            close_img = ctk.CTkImage(Image.open(self.config['Close']['name']), size=(80, 80))
            buttons_dict['Exit'] = ctk.CTkButton(
                master=self.window,
                text="",
                command=self.window.destroy,
                fg_color='transparent',
                bg_color='white',
                hover_color='#F2F2F2',
                image=close_img
            )
            buttons_dict['Exit'].place(x=self.config['Close']['x'], y=self.config['Close']['y'])
        
        self.buttons = buttons_dict

    # --- Eingabefelder (Entries) erstellen ---
    def create_entries(self, tfh_obj):
        """
        Erstellt Eingabefelder für verschiedene Gerätetypen (mfc, Vorgabe, Modbus_Pump).
        
        Die Positionierung erfolgt über .grid() oder .place() basierend auf der Konfiguration.
        """
        entries_dict = {
            'mfc': {},
            'Vorgabe': {},
            'ExtOutput': {},
            'Modbus_Pump': {}
        }
        i_MFC, i_V, i_MP = 0, 0, 0
        index_counters = {key: 0 for key in entries_dict}

        # Erzeuge Eingabefelder für mfc-Geräte aus der Modbus-Konfiguration
        for control_name, control_rule in self.modbus_obj.config.items():
            if control_rule.get("type") == "mfc":
                parent_var, options = (
                    (self.frames['mfc'], {'grid_opts': {'column': 2, 'row': i_MFC + 1, 'ipadx': 7, 'ipady': 7, 'padx': 2}})
                    if control_rule.get("Box") == 1
                    else (self.window, {'x': control_rule.get("x") + 25, 'y': control_rule.get("y") + 45})
                )
                entries_dict['mfc'][i_MFC] = self._create_entry(
                    parent=parent_var,
                    default_text="0",
                    font=('Arial', 18),
                    width=40,
                    fg_color='light blue',
                    **options
                )
                entries_dict['mfc'][i_MFC].deviceName = control_name
                i_MFC += 1
            if control_rule.get("type") == "ExtOutput":
                ic = index_counters['ExtOutput']
                parent_var, options = (
                    (self.frames['mfc'], {'grid_opts': {'column': 2, 'row': ic + 1, 'ipadx': 7, 'ipady': 7, 'padx': 2}})
                    if control_rule.get("Box") == 1
                    else (self.window, {'x': control_rule.get("x") + 25, 'y': control_rule.get("y") + 45})
                )
                entries_dict['ExtOutput'][ic] = self._create_entry(
                    parent=parent_var,
                    default_text="0",
                    font=('Arial', 18),
                    width=40,
                    fg_color='light blue',
                    **options
                )
                entries_dict['ExtOutput'][ic].deviceName = control_name
                index_counters['ExtOutput'] += 1

        # Erzeuge weitere Eingabefelder anhand der tfh-Konfiguration
        for control_name, control_rule in tfh_obj.config.items():
            device_type = control_rule.get("type")
            if device_type == "mfc" and "modbus" not in control_rule.get("input_device", "").lower():
                entries_dict['mfc'][i_MFC] = self._create_entry(
                    parent=self.window,
                    default_text="0",
                    x=control_rule.get("x"),
                    y=control_rule.get("y"),
                    font=('Arial', 18),
                    width=40,
                    fg_color='light blue'
                )
                entries_dict['mfc'][i_MFC].deviceName = control_name
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
                entries_dict['Vorgabe'][i_V].deviceName = control_name
                i_V += 1
            elif device_type == "Modbus_Pump":
                entries_dict['Modbus_Pump'][i_MP] = self._create_entry(
                    parent=self.window,
                    default_text="0",
                    x=control_rule.get("x"),
                    y=control_rule.get("y"),
                    font=('Arial', 18),
                    width=40,
                    fg_color='light blue'
                )
                entries_dict['Vorgabe'][i_MP].deviceName = control_name
                i_MP += 1

        # Speichere Standard-Dateipfade für Save/Excel-Funktion
        entries_dict['SaveFile'] = "../Daten/test.dat"
        entries_dict['ExcelFile'] = "../Excel.xlsx"
        self.entries = entries_dict

    # --- Controller (z. B. easy_PI) einrichten ---
    def setup_controller(self, tfh_obj):
        """
        Richtet Controller basierend auf der Konfiguration ein.
        
        Für jeden Controller wird neben dem Regelungsobjekt auch ein Eingabefeld (Entry)
        und ein Label zur Anzeige des Ausgangswerts erstellt.
        """
        controllers_dict = {'easy_PI': {},'direct_Heat': {}}
        i_PI = 0
        i_directHeat = 0

        for control_name, control_rule in tfh_obj.config.items():
            if control_rule.get("type") == "direct_Heat": # Keine Regelung sondern direkte Vorgabe der %-tualen Heizleistung
                
                out_device = control_rule.get("output_device")
                out_channel = control_rule.get("output_channel")
                # Speichere den control_name als Attribut
                controllers_dict['direct_Heat'][i_directHeat] = DirectHeatController(control_name)
                # Erzeuge Eingabefeld für den Vorgabewert
                controllers_dict['direct_Heat'][i_directHeat].entry = ctk.CTkEntry(
                    self.window,
                    font=('Arial', 16),
                    width=50,
                    fg_color='light blue'
                )
                controllers_dict['direct_Heat'][i_directHeat].entry.place(x=control_rule.get("x"), y=control_rule.get("y"))
                
                # Erzeuge Label zur Anzeige des Ausgangswerts
                controllers_dict['direct_Heat'][i_directHeat].label = ctk.CTkLabel(
                    self.window,
                    font=('Arial', 18),
                    text='0 %',
                    bg_color='white'
                )
                controllers_dict['direct_Heat'][i_directHeat].label.place(x=control_rule.get("x"), y=control_rule.get("y") + 35)

                i_directHeat += 1



            if control_rule.get("type") == "easy_PI":
                out_device = control_rule.get("output_device")
                out_channel = control_rule.get("output_channel")
                P_val = control_rule["DeviceInfo"].get("P_Value")
                I_val = control_rule["DeviceInfo"].get("I_Value")

                # Wähle den Eingang: extern oder über ein anderes Gerät
                if "extern" in control_rule.get("input_device", "").lower():
                    controllers_dict['easy_PI'][i_PI] = easy_PI(out_device, out_channel, "extern", 0, I_val, P_val)
                else:
                    in_device = tfh_obj.config[control_rule.get("input_device")].get("input_device")
                    controllers_dict['easy_PI'][i_PI] = easy_PI(out_device, out_channel, tfh_obj.inputs[in_device], 0, I_val, P_val)
                
                # Speichere den control_name als Attribut
                controllers_dict['easy_PI'][i_PI].deviceName = control_name
                # Erzeuge Eingabefeld für den Sollwert
                controllers_dict['easy_PI'][i_PI].entry = ctk.CTkEntry(
                    self.window,
                    font=('Arial', 16),
                    width=50,
                    fg_color='light blue'
                )
                controllers_dict['easy_PI'][i_PI].entry.place(x=control_rule.get("x"), y=control_rule.get("y"))

                # Erzeuge Label zur Anzeige des Ausgangswerts
                controllers_dict['easy_PI'][i_PI].label = ctk.CTkLabel(
                    self.window,
                    font=('Arial', 18),
                    text='0 %',
                    bg_color='white'
                )
                controllers_dict['easy_PI'][i_PI].label.place(x=control_rule.get("x"), y=control_rule.get("y") + 35)

                i_PI += 1

        self.controller = controllers_dict

    # --- Werte in Datei speichern ---
    def save_values(self):
        """
        Schreibt aktuelle Werte der Geräte in eine Logdatei.
        
        Beim ersten Aufruf wird ein Header mit Geräteinformationen und Spaltenüberschriften geschrieben.
        Danach wird in regelmäßigen Abständen (alle ca. 1 Sekunde) eine Zeile mit Zeitstempel
        und den Mess-/Eingabewerten angehängt.
        """
        if self.write_header:
            write_device_informations(self, self.tfh_obj)
            header_comment = "### Device Names"
            header_columns = ["Zeitpunkt"]
            
            for control_name, control_rule in self.modbus_obj.config.items():  
                device_type = control_rule.get("type")
                if device_type == "mfc":
                    header_columns.extend([f"{control_name}_Soll", f"{control_name}_Ist"])
                else:
                    header_columns.append(control_name)

            for control_name, control_rule in self.tfh_obj.config.items():
                device_type = control_rule.get("type")
                if device_type == "mfc":
                    header_columns.extend([f"{control_name}_Soll", f"{control_name}_Ist"])
                elif device_type == "easy_PI":
                    header_columns.extend([f"{control_name}_Soll", f"{control_name}_Output"])
                elif device_type == "direct_Heat":
                    header_columns.extend([f"{control_name}_Soll", f"{control_name}_Output"])
                else:
                    header_columns.append(control_name)
            with open(self.entries['SaveFile'], 'a') as f:
                f.write(header_comment + "\n" + "\t".join(header_columns) + "\n")
            self.write_header = False

        i_MFC, i_PI, i_V, i_MP,i_directHeat = 0, 0, 0, 0, 0
        data_columns = [datetime.now().strftime('%Y-%m-%d %H:%M:%S.%f')]
        
           
            
        for control_name, control_rule in self.modbus_obj.config.items():            
            if control_rule.get("type") == "mfc":
                if self.modbus_obj.operation_mode != 1:  
                    mfc_value = self.entries['mfc'][i_MFC].get()
                    if self.tfh_obj.operation_mode != 1:
                        input_val = self.modbus_obj.devices[control_name].flow
                    else:
                        input_val = 0.0
                    data_columns.extend([str(mfc_value), str(input_val)])
                    i_MFC += 1
            elif control_rule.get("type") == "ExtOutput":
                entry_id = self.getID("ExtOutput", control_name)
                output_value = self.entries['ExtOutput'][entry_id].get()
                data_columns.append(str(output_value))

        for control_name, control_rule in self.tfh_obj.config.items():
            device_type = control_rule.get("type")
            input_device_uid = control_rule.get("input_device")
            input_channel = control_rule.get("input_channel")
            output_device_uid = control_rule.get("output_device")
            output_channel = control_rule.get("output_channel")

            if device_type in ("thermocouple", "pressure", "FlowMeter", "ExtInput", "analytic"):
                input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                data_columns.append(str(input_val))
            elif device_type == "valve":
                output_val = int(self.tfh_obj.outputs[output_device_uid].values[output_channel])
                data_columns.append(str(output_val))
            elif device_type == "Vorgabe":
                vorgabe_value = self.entries['Vorgabe'][i_V].get()
                data_columns.append(str(vorgabe_value))
                i_V += 1
            elif device_type == "Modbus_Pump":
                vorgabe_value = self.entries['Modbus_Pump'][i_MP].get()
                data_columns.append(str(vorgabe_value))
                i_MP += 1
            elif device_type == "mfc":
                mfc_value = self.entries['mfc'][i_MFC].get()
                if self.tfh_obj.operation_mode != 1:
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                else:
                    input_val = 0.0
                data_columns.extend([str(mfc_value), str(input_val)])
                i_MFC += 1
            elif device_type == "easy_PI":
                output_percent = self.controller['easy_PI'][i_PI].out * 100
                setpoint_value = self.controller['easy_PI'][i_PI].soll
                data_columns.extend([str(setpoint_value), str(output_percent)])
                i_PI += 1
            elif device_type == "direct_Heat":
                output_percent = self.controller['direct_Heat'][i_directHeat].out * 100
                setpoint_value = self.controller['direct_Heat'][i_directHeat].soll
                data_columns.extend([str(setpoint_value), str(output_percent)])
                i_directHeat += 1

        with open(self.entries['SaveFile'], 'a') as f:
            f.write("\t".join(data_columns) + "\n")

    # --- Excel-Funktionen ---
    def start_excel(self):
        """
        Startet den Excel-Modus:
         - Lädt die Excel-Datei
         - Setzt den Startpunkt und initialisiert den Timer
         - Deaktiviert den Start-Button und aktiviert den Save-Switch
        """
        self.excel_data = openpyxl.load_workbook(self.entries['ExcelFile'], data_only=True)
        self.running_excel = 1
        self.section = 4  # Start in Zeile 4
        self.t0 = time.time()
        self.sheet = self.excel_data["Ablauf"]
        tmp = self.sheet[1][1]
        if tmp.value is None:
            messagebox.showerror("Excelfehler", "Laufzeit nicht in Excelsheet!")
        self.run_time = tmp.value * 60 + self.t0
        self.buttons['Save'].select()
        self.buttons['StartExcel'].configure(state="disabled")
        print("Start")

    def stop_excel(self):
        """
        Stoppt den Excel-Modus, reaktiviert den Start-Button und setzt den Timer zurück.
        """
        self.running_excel = 0
        self.buttons['StartExcel'].configure(state="enabled")
        self.buttons['Save'].deselect()
        self.labels["Timer"].configure(text="0.00 min")
        print("Stop")

    # --- Werte an die Geräte senden ---
    def set_data(self):
        """
        Liest die Eingabefelder und Controller-Eingaben aus und schreibt die Werte an die entsprechenden Geräte.
        """
        i_MFC, i_PI, i_MP,i_directHeat = 0, 0, 0, 0
        entries = self.entries
        controller = self.controller
        modbus_obj = self.modbus_obj
        entries_dict = {
            'mfc': {},
            'Vorgabe': {},
            'ExtOutput': {},
            'Modbus_Pump': {}
        }
        index_counters = {key: 0 for key in entries_dict}
           
            
        for control_name, control_rule in self.modbus_obj.config.items():
            unit = control_rule["DeviceInfo"].get("unit")
            
            if control_rule.get("type") == "mfc":
                if self.modbus_obj.operation_mode != 1:                    
                    value = float(entries['mfc'][i_MFC].get())
                    self.modbus_obj.devices[control_name].set(value)
                i_MFC += 1

        
            if control_rule.get("type") == "ExtOutput":
                if self.modbus_obj.operation_mode != 1:    
                    ic = index_counters['ExtOutput']                    
                    value = float(entries['ExtOutput'][ic].get())
                    self.modbus_obj.devices[control_name].set(value)
               
                index_counters['ExtOutput'] += 1




        for control_name, control_rule in self.tfh_obj.config.items():
            device_type = control_rule.get("type")
            output_channel = control_rule.get("output_channel")
            output_device_uid = control_rule.get("output_device")
            gradient = control_rule["DeviceInfo"].get("gradient")
            y_axis = control_rule["DeviceInfo"].get("y-axis")
            
            if device_type == "mfc":
                if entries['mfc'][i_MFC].get() != '':
                    value = float(entries['mfc'][i_MFC].get()) / gradient + y_axis
                    if self.tfh_obj.operation_mode != 1:
                        self.tfh_obj.outputs[output_device_uid].values[output_channel] = value
                i_MFC += 1

            if device_type == "Modbus_Pump":
                if entries['Modbus_Pump'][i_MP].get() != '':
                    value = float(entries['Modbus_Pump'][i_MP].get())
                    if self.tfh_obj.operation_mode != 1:
                        device = modbus_obj[i_MP]
                        device.set_Flow(value, gradient, y_axis)
                i_MP += 1

            if device_type == "easy_PI":
                if controller['easy_PI'][i_PI].entry.get() != '':
                    value = float(controller['easy_PI'][i_PI].entry.get())
                    if not controller['easy_PI'][i_PI].running:
                        controller['easy_PI'][i_PI].start(value)
                    else:
                        controller['easy_PI'][i_PI].set_soll(value)
                i_PI += 1

            if device_type == "direct_Heat":
                if controller['direct_Heat'][i_directHeat].entry.get() != '':
                    value = float(controller['direct_Heat'][i_directHeat].entry.get())
                    if not controller['direct_Heat'][i_directHeat].running:
                        controller['direct_Heat'][i_directHeat].start(value)
                    else:
                        controller['direct_Heat'][i_directHeat].set_soll(value)
                i_directHeat += 1

    # --- Dateiauswahlfunktionen ---
    def get_file(self):
        """
        Öffnet einen Dialog zur Dateiauswahl für den Speicherpfad und aktualisiert den entsprechenden Eintrag.
        """
        file_path = asksaveasfilename(defaultextension=".dat", initialdir="./Daten/")
        if file_path:
            self.entries['SaveFile'] = file_path
            parent_folder = os.path.basename(os.path.dirname(file_path))
            file_name = os.path.basename(file_path)
            short_text = f"{parent_folder}/{file_name}"
            if 'Save' in self.labels:
                self.labels['Save'].configure(text=short_text)

    def get_Excelfile(self):
        """
        Öffnet einen Dialog zur Auswahl einer Excel-Datei und aktualisiert den entsprechenden Eintrag.
        """
        file_path = askopenfilename(defaultextension=".xlsx", initialdir="./")
        if file_path:
            self.entries['ExcelFile'] = file_path
            parent_folder = os.path.basename(os.path.dirname(file_path))
            file_name = os.path.basename(file_path)
            short_text = f"{parent_folder}/{file_name}"
            if 'ExcelFile' in self.labels:
                self.labels['ExcelFile'].configure(text=short_text)

    # --- Hauptschleife ---
    def run(self):
        """Startet die Hauptschleife der GUI."""
        self.window.mainloop()


    def getID(self, ctrl_type, device_name):
        """
        Sucht in self.controller nach einem Controller des Typs ctrl_type,
        dessen deviceName dem übergebenen device_name entspricht.
        Gibt den Index zurück oder None, falls nicht gefunden.
        """
        
        if ctrl_type == "easy_PI":
            for idx, controller in self.controller.get(ctrl_type, {}).items():
                if controller.deviceName == device_name:
                    return idx
        if ctrl_type == "direct_Heat":
            for idx, controller in self.controller.get(ctrl_type, {}).items():
                if controller.deviceName == device_name:
                    return idx
        if ctrl_type == "mfc" or ctrl_type == "ExtOutput":
            for idx, entry in self.entries.get(ctrl_type, {}).items():
                if entry.deviceName == device_name:
                    return idx
        return None


    def start_loop(self):
        """
        Aktualisiert periodisch die GUI:
          - Verarbeitet Excel-Daten, falls der Excel-Modus aktiv ist.
          - Aktualisiert Labels und Controller anhand der Sensordaten.
          - Ruft save_values() periodisch auf.
          - Plant den nächsten Aufruf in 50ms.
        """
        i_MFC, i_Tc, i_PI, i_p,i_a, i_exI, i_FI,i_directHeat = 0, 0, 0, 0, 0, 0, 0, 0
        
        # Excel-Modus: Aktualisiere Timer und Eingaben aus Excel
        if self.running_excel == 1:
            entries = self.entries
            controller = self.controller
            self.t_end = self.run_time - time.time()
            output, self.section, self.t_section, self.t0 = Excel_timing(self.sheet, self.section, self.t0)
            self.labels['Timer'].configure(text=f"{self.t_end/60:.2f} min")
            for control_name, control_rule in self.modbus_obj.config.items():
                # Aktualisiere den Controller und die mfc-Eingaben
                if control_rule.get("type") == "easy_PI":
                    controller_id = self.getID("easy_PI", control_name)
                    controller['easy_PI'][controller_id].entry.delete(0, tk.END)
                    controller['easy_PI'][controller_id].entry.insert(0, f"{output[control_name]:.2f}")
               
                elif control_rule.get("type") == "mfc":
                    mfc_id = self.getID("mfc", control_name)
                    entries['mfc'][mfc_id].delete(0, tk.END)
                    entries['mfc'][mfc_id].insert(0, f"{output[control_name]:.2f}")
                elif control_rule.get("type") == "ExtOutput":
                    entry_id = self.getID("ExtOutput", control_name)
                    entries['ExtOutput'][entry_id].delete(0, tk.END)
                    entries['ExtOutput'][entry_id].insert(0, f"{output[control_name]:.2f}")
                    
            
            for control_name, control_rule in self.tfh_obj.config.items():
                # Aktualisiere den Controller und die mfc-Eingaben
                if control_rule.get("type") == "easy_PI":
                    controller_id = self.getID("easy_PI", control_name)
                    controller['easy_PI'][controller_id].entry.delete(0, tk.END)
                    controller['easy_PI'][controller_id].entry.insert(0, f"{output[control_name]:.2f}")
                elif control_rule.get("type") == "direct_Heat":
                    controller_id = self.getID("direct_Heat", control_name)
                    controller['direct_Heat'][controller_id].entry.delete(0, tk.END)
                    controller['direct_Heat'][controller_id].entry.insert(0, f"{output[control_name]:.2f}")
                elif control_rule.get("type") == "mfc":
                    mfc_id = self.getID("mfc", control_name)
                    entries['mfc'][mfc_id].delete(0, tk.END)
                    entries['mfc'][mfc_id].insert(0, f"{output[control_name]:.2f}")
                elif control_rule.get("type") == "valve":
                    if output[control_name] == 1:
                        self.buttons[control_name].select() 
                    else:
                        self.buttons[control_name].deselect() 
                                    
            self.set_data()
                
            if self.t_end < 0:
                self.stop_excel()
        
        # Aktualisiere die Labels basierend auf den aktuellen Sensordaten
        for control_name, control_rule in self.modbus_obj.config.items():
            unit = control_rule["DeviceInfo"].get("unit")
            
            if control_rule.get("type") == "mfc":
                if self.modbus_obj.operation_mode != 1:                    
                    value = self.modbus_obj.devices[control_name].flow
                if value is not None:
                    text = f"{round(value, 0)} {unit}"
                else:
                    text = "Error"  # oder ein anderer Platzhalter/Text
                self.labels['mfc'][i_MFC].configure(text=text)
                i_MFC += 1


        for control_name, control_rule in self.tfh_obj.config.items():
            input_channel = control_rule.get("input_channel")
            input_device_uid = control_rule.get("input_device")
            output_channel = control_rule.get("output_channel")
            output_device_uid = control_rule.get("output_device")
            gradient = control_rule["DeviceInfo"].get("gradient")
            y_axis = control_rule["DeviceInfo"].get("y-axis")
            unit = control_rule["DeviceInfo"].get("unit")
            device_type = control_rule.get("type")
            
            if device_type == "thermocouple":
                input_val = self.tfh_obj.inputs[input_device_uid].values[0]
                self.labels['Tc'][i_Tc].configure(text=f"{round(input_val, 2)} {unit}")
                i_Tc += 1

            elif device_type == "pressure":
                if self.tfh_obj.operation_mode != 1:
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    converted_value = ((input_val/1e6 - y_axis) * gradient)
                    converted_value = input_val
                    self.labels['Pressure'][i_p].configure(text=f"{round(converted_value, 2)} {unit}")
                i_p += 1

            elif device_type == "analytic":
                if self.tfh_obj.operation_mode != 1:
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    converted_value = ((input_val/1e6 - y_axis) * gradient)
                    converted_value = input_val
                    #print(f"reading input on device {control_name} - {input_channel} {input_val}")
                    self.labels['analytic'][i_a].configure(text=f"{round(converted_value, 2)} {unit}")
                i_a += 1

            elif device_type == "FlowMeter":
                if self.tfh_obj.operation_mode != 1:
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    converted_value = 0 + (100 - 0) / (20 - 4) * (input_val/1e6 - 4)
                    converted_value = max(converted_value, 0)
                    self.labels['FlowMeter'][i_FI].configure(text=f"{round(converted_value, 2)} {unit}")
                i_FI += 1

            elif device_type == "mfc":
                if self.tfh_obj.operation_mode != 1:
                    input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                    converted_value = (input_val - y_axis) * gradient
                    self.labels['mfc'][i_MFC].configure(text=f"{round(converted_value, 2)} {unit}")
                i_MFC += 1

            elif device_type == "easy_PI":
                
                if  self.controller['direct_Heat'][0].out <= 0:
                    self.controller['easy_PI'][i_PI].regeln()

                value = self.controller['easy_PI'][i_PI].out
                if control_rule["DeviceInfo"].get('Power', False):
                    Power = control_rule["DeviceInfo"].get('Power')
                    self.controller['easy_PI'][i_PI].label.configure(text=f"{value*Power:.2f} {unit}")
                if control_rule.get("output_type") == "analog_mA":
                    value = (4 + (20 - 4) * value) * 1000
                self.tfh_obj.outputs[output_device_uid].values[output_channel] = value
                i_PI += 1

            elif device_type == "direct_Heat":
                value = self.controller['direct_Heat'][i_directHeat].out/100 # Vorgabe in Prozent
                if control_rule["DeviceInfo"].get('Power', False):
                    Power = control_rule["DeviceInfo"].get('Power')
                    self.controller['direct_Heat'][i_directHeat].label.configure(text=f"{value*Power:.2f} {unit}")
                self.tfh_obj.outputs[output_device_uid].values[output_channel] = value
                i_directHeat += 1

            elif device_type == "ExtInput":
                input_val = self.tfh_obj.inputs[input_device_uid].values[input_channel]
                self.labels['ExtInput'][i_exI].configure(text=f"{round(input_val / 1e6, 2)} mA")
                i_exI += 1
                if control_rule["DeviceInfo"].get('Power', False):
                    converted_value = (input_val - y_axis) * gradient
                    converted_value = max(converted_value, 0)
                    self.labels['ExtInput'][i_exI].configure(text=f"{round(converted_value, 2)} {unit}")
                i_exI += 1

            elif device_type == "valve":
                if self.buttons[control_name].get() == 1:
                    self.tfh_obj.outputs[output_device_uid].values[output_channel] = True
                else:
                    self.tfh_obj.outputs[output_device_uid].values[output_channel] = False

        # Speichere Werte, wenn der Save-Switch aktiv ist und mehr als 1 Sekunde vergangen ist
        if self.buttons['Save'].get() == 1 and time.time() - self.save_timer > 1:
            self.save_values()
            self.save_timer = time.time()

        # Plane den nächsten Aufruf in 50 ms
        self.window.after(50, self.start_loop)

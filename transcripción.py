# encoding: utf-8

import ConfigParser
import os
import sys
import Tkinter, tkFileDialog
from Tkinter import *

from openpyxl import load_workbook

__version__ = "1.0"


CONFIG_FILE = u'transcripción.cfg'
XLSX_PREFIX = u'Padrones electorales 1925 Santiago'
tkinter_umlauts=['odiaeresis', 'adiaeresis', 'udiaeresis', 'Odiaeresis', 'Adiaeresis', 'Udiaeresis', 'ssharp']


class AutocompleteEntry(Tkinter.Entry):
        
        def get_completion_list(self): return self._completion_list

        def set_completion_list(self, completion_list):
                self._completion_list = completion_list
                self._hits = []
                self._hit_index = 0
                self.position = 0
                self.bind('<KeyRelease>', self.handle_keyrelease)               

        def autocomplete(self, delta=0):
                """autocomplete the Entry, delta may be 0/1/-1 to cycle through possible hits"""
                if delta: # need to delete selection otherwise we would fix the current position
                        self.delete(self.position, Tkinter.END)
                else: # set position to end so selection starts where textentry ended
                        self.position = len(self.get())
                # collect hits
                _hits = []
                for element in self._completion_list:
                        if element.startswith(self.get().lower()):
                                _hits.append(element)
                # if we have a new hit list, keep this in mind
                if _hits != self._hits:
                        self._hit_index = 0
                        self._hits=_hits
                # only allow cycling if we are in a known hit list
                if _hits == self._hits and self._hits:
                        self._hit_index = (self._hit_index + delta) % len(self._hits)
                # now finally perform the auto completion
                if self._hits:
                        self.delete(0,Tkinter.END)
                        self.insert(0,self._hits[self._hit_index])
                        self.select_range(self.position,Tkinter.END)
                        
        def handle_keyrelease(self, event):
                """event handler for the keyrelease event on this widget"""
                if event.keysym == "BackSpace":
                        self.delete(self.index(Tkinter.INSERT), Tkinter.END) 
                        self.position = self.index(Tkinter.END)
                if event.keysym == "Left":
                        if self.position < self.index(Tkinter.END): # delete the selection
                                self.delete(self.position, Tkinter.END)
                        else:
                                self.position = self.position-1 # delete one character
                                self.delete(self.position, Tkinter.END)
                if event.keysym == "Right":
                        self.position = self.index(Tkinter.END) # go to end (no selection)
                if event.keysym == "Down":
                        self.autocomplete(1) # cycle to next hit
                if event.keysym == "Up":
                        self.autocomplete(-1) # cycle to previous hit
                # perform normal autocomplete if event is a single key or an umlaut
                if len(event.keysym) == 1 or event.keysym in tkinter_umlauts:
                        self.autocomplete()


class Transciption(Tkinter.Tk):
        def __init__(self, *args, **kwargs):
                Tkinter.Tk.__init__(self, *args, **kwargs)

                self.title(u' Gaby\'s Transcription :) :)')

                self.init_values()
                
                current_frame = Frame(self)
                current_frame.pack()
                Label(current_frame, text=u'Fila actual :  ').pack(side=LEFT)
                self.var_current_row = StringVar()
                Label(current_frame, textvariable=self.var_current_row).pack()


                self.read_config()
                self.add_fields()
                
                options_frame = Frame(self)
                options_frame.pack()
                self.boton=Button(options_frame,text="Primero", command=self.first)
                self.boton.pack(side=LEFT)
                self.boton=Button(options_frame,text="Cargar planilla", command=self.load)
                self.boton.pack(side=LEFT)
                self.boton=Button(options_frame,text="Último", command=self.last)
                self.boton.pack()

                options_frame = Frame(self)
                options_frame.pack()
                self.boton=Button(options_frame,text="Anterior", command=self.previuos)
                self.boton.pack(side=LEFT)
                self.boton=Button(options_frame,text="Guardar", command=self.save)
                self.boton.pack(side=LEFT)
                self.boton=Button(options_frame,text="Siguiente", command=self.next)
                self.boton.pack()

        def init_values(self):
                self.config = None
                self.current_row = None
                self.xlsx_name = None
                self.wb = None
                self.ws = None


        def read_config(self):
                config = dict()
                parser = ConfigParser.RawConfigParser()
                parser.read(CONFIG_FILE)
                sections = parser.sections()
                for section in sections:
                        config[section] = dict()
                        for option in parser.options(sections[0]):
                                if option == 'autocomplete':
                                        config[section]['autocomplete'] = parser.get(section, 'autocomplete')
                                        if config[section]['autocomplete']:
                                                config[section]['autocomplete'] = config[section]['autocomplete'].split(' ')
                                                config[section]['autocomplete'].append('')
                                        else:
                                                config[section]['autocomplete'] = ['']
                                        config[section]['autocomplete'] = [text.replace('_', ' ') for text in config[section]['autocomplete']]
                                else:
                                        config[section][option] = parser.get(section, option)
                self.config = config

        def add_fields(self):
                self.fields = {}
                self.vars_fields = {}
                for field in self.config:
                        frame = Frame(self)
                        frame.pack()
                        Label(frame, text=field, anchor=E, width=35).pack(fill=X, side=LEFT)
                        self.vars_fields[field] = StringVar()
                        self.fields[field] = AutocompleteEntry(frame, textvariable=self.vars_fields[field])
                        self.fields[field].set_completion_list(self.config[field]['autocomplete'])
                        self.fields[field].pack()

        def clean_entries(self):
                for field in self.fields:
                        self.vars_fields[field].set('')

        def show_cell(self, row):
                print 'Showing' , row

                self.var_current_row.set(str(row))
                # Save current worksheet (not persist)
                self.save_ws()

                self.current_row = row

                # Show on Entries
                self.clean_entries()
                for field in self.config:
                        column = self.config[field].get('column')
                        if column:
                                value = self.ws.cell(column + str(row)).value
                                if value:
                                        self.vars_fields[field].set(value)
                print 'Showing OK'

        def load(self):
                print 'Openning ' , self.xlsx_name
                self.xlsx_name = tkFileDialog.askopenfilename(parent=self, title='Escoge la planilla para trabajar', defaultextension='xlsx')
                if not self.xlsx_name: return
                self.wb = load_workbook(self.xlsx_name)
                self.ws = self.wb.worksheets[0]

                last_row = self.ws.get_highest_row()
                self.show_cell(last_row)
                print 'Openning OK'

        def save(self):
                self.save_ws(persist=True)
        
        def save_ws(self, persist=False):
                print 'Saving Current Row ...'
                if self.current_row:
                        # Respaldar los actuales datos si no existe archivo
                        if not self.xlsx_name:
                                current_data = dict()
                                for field in self.fields:
                                        current_data[field] = self.vars_fields[field].get()
                                self.load()
                                self.next()
                                for field in self.fields:
                                        self.vars_fields[field].set(current_data[field])

                        # Guardar nuevo dato en el XLSX
                        for field in self.config:
                                column = self.config[field].get('column')
                                if column:
                                        cell = self.ws.cell(column + str(self.current_row))
                                        cell.value = self.vars_fields[field].get()
                        if persist:
                                self.wb.save(self.xlsx_name)
                print 'Saving OK'

        def update_autocomplete(self):
                # Actualizar la lista de autocompletado en los Entry
                new_texts = dict()
                for field in self.fields:
                        text = self.fields[field].get()
                        if text:
                                text = text.replace(' ', '_')
                                completion_list = self.fields[field].get_completion_list()
                                if not text in completion_list:
                                        new_texts[field] = text
                                        completion_list.append(text)
                                        self.fields[field].set_completion_list(completion_list)

                # Actualizar el archivo de conf con la lista de autocompletado
                if new_texts.items():
                        parser = ConfigParser.RawConfigParser()
                        parser.read(CONFIG_FILE)
                        for field in new_texts:
                                autocomplete_value = parser.get(field, 'autocomplete')
                                autocomplete_value += ' ' + new_texts[field]
                                parser.set(field, 'autocomplete', autocomplete_value)
                        with open(CONFIG_FILE, 'wb') as configfile:
                            parser.write(configfile)
                        self.read_config()


        def first(self):
                self.show_cell(2)

        def last(self):
                last_row = self.ws.get_highest_row()
                self.show_cell(last_row)

        def previuos(self):
                if self.current_row > 2:
                        self.show_cell(self.current_row-1)

        def is_current_cell_not_empty(self):
                empty = True
                for field in self.vars_fields:
                        if self.vars_fields[field].get():
                                empty = False
                                break
                return not empty

        def next(self):
                if self.is_current_cell_not_empty():
                        self.show_cell(self.current_row+1)


if __name__ == '__main__':
        t = Transciption()
        t.mainloop()
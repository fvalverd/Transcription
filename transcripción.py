#!/usr/bin/env python
# encoding: utf-8

import sys
import os
import Tkinter
from Tkinter import *
import ConfigParser

__version__ = "1.0"

tkinter_umlauts=['odiaeresis', 'adiaeresis', 'udiaeresis', 'Odiaeresis', 'Adiaeresis', 'Udiaeresis', 'ssharp']


class AutocompleteEntry(Tkinter.Entry):
        
        def get_completion_list(self):
                return self._completion_list

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

                self.title(u' Gaby\'s Transcription :)')

                config = self.read_config()
                self.add_fields(config)
                
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

        def read_config(self):
                config = dict()
                parser = ConfigParser.RawConfigParser()
                parser.read(u'transcripción.cfg')
                sections = parser.sections()
                for section in sections:
                        config[section] = dict()
                        config[section]['autocomplete'] = parser.get(section, 'autocomplete')
                        if config[section]['autocomplete']:
                                config[section]['autocomplete'] = config[section]['autocomplete'].split(' ')
                                config[section]['autocomplete'].append('')
                        else:
                                config[section]['autocomplete'] = ['']
                        config[section]['autocomplete'] = [text.replace('_', ' ') for text in config[section]['autocomplete']]
                return config

        def add_fields(self, config):
                self.fields = {}
                for field in config:
                        frame = Frame(self)
                        frame.pack()
                        Label(frame, text=field, anchor=E, width=35).pack(fill=X, side=LEFT)
                        self.fields[field] = AutocompleteEntry(frame)
                        self.fields[field].set_completion_list(config[field]['autocomplete'])
                        self.fields[field].pack()

                # self.fields[config[u'Inscripción']].focus_set()


        def load(self) : pass
        def save(self) :
                # TODO: generar el XLSX

                # TODO: Actualizar el archivo de conf con la lista de autocompletado
                # Revisar que el autocompletado con espacio se cambie por _

                # Actualizar la lista de autocompletado
                for field in self.fields:
                        text = self.fields[field].get()
                        if text:
                                completion_list = self.fields[field].get_completion_list()
                                if not text in completion_list:
                                        completion_list.append(text)
                                        self.fields[field].set_completion_list(completion_list)

        def first(self) : pass
        def last(self) : pass
        def previuos(self) : pass
        def next(self) : pass


if __name__ == '__main__':
        t = Transciption()
        t.mainloop()
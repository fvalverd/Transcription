#!/usr/bin/env python
# encoding: utf-8

import sys
import os
import Tkinter
from Tkinter import *

__version__ = "1.0"

tkinter_umlauts=['odiaeresis', 'adiaeresis', 'udiaeresis', 'Odiaeresis', 'Adiaeresis', 'Udiaeresis', 'ssharp']


class AutocompleteEntry(Tkinter.Entry):
        
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

def run_transcription():
        test_list = (u'test', u'type', u'true', u'tree', u'tölz')
        root = Tkinter.Tk(className=u' Gaby\'s Transcription :)')

        # A # Inscripción
        registration_frame = Frame(root)
        registration_frame.pack()

        label = Label(registration_frame, text=u'Inscripción')
        label.pack(side=LEFT)

        registration_list = (u'extraordinaria', u'')
        registration_entry = AutocompleteEntry(registration_frame)
        registration_entry.set_completion_list(registration_list)
        registration_entry.pack()
        registration_entry.focus_set()

        
        # C # Sección
        section_frame = Frame(root)
        section_frame.pack()

        label = Label(section_frame, text=u'N° Sección')
        label.pack(side=LEFT)

        section_entry = AutocompleteEntry(section_frame)
        section_entry.set_completion_list(test_list)
        section_entry.pack()
        section_entry.focus_set()


        # F # Comuna Subdelegación
        commune_subbranch_frame = Frame(root)
        commune_subbranch_frame.pack()

        label = Label(commune_subbranch_frame, text=u'Comuna Subdelegación')
        label.pack(side=LEFT)

        commune_subbranch_list = (u'sta lucia', u'')
        commune_subbranch_entry = AutocompleteEntry(commune_subbranch_frame)
        commune_subbranch_entry.set_completion_list(commune_subbranch_list)
        commune_subbranch_entry.pack()
        commune_subbranch_entry.focus_set()  


        # G # Subdelegación
        subbranch_frame = Frame(root)
        subbranch_frame.pack()

        label = Label(subbranch_frame, text=u'N° Subdelegación')
        label.pack(side=LEFT)

        subbranch_list = (u'sta lucia', u'')
        subbranch_entry = AutocompleteEntry(subbranch_frame)
        subbranch_entry.set_completion_list(subbranch_list)
        subbranch_entry.pack()
        subbranch_entry.focus_set()


        # I # N° Inscripción
        n_registration_frame = Frame(root)
        n_registration_frame.pack()

        label = Label(n_registration_frame, text=u'N° Inscripción')
        label.pack(side=LEFT)

        n_registration_entry = Entry(n_registration_frame)
        n_registration_entry.pack()
        n_registration_entry.focus_set()



        # K # Extranjero
        foreign_frame = Frame(root)
        foreign_frame.pack()

        label = Label(foreign_frame, text=u'Extranjero')
        label.pack(side=LEFT)

        foreign_list = (u'???', u'')
        foreign_entry = AutocompleteEntry(foreign_frame)
        foreign_entry.set_completion_list(foreign_list)
        foreign_entry.pack()
        foreign_entry.focus_set()


        # L # Gabinete que otorgó el carnet de identidad
        place_frame = Frame(root)
        place_frame.pack()

        label = Label(place_frame, text=u'Gabinete que otorgó el carnet de identidad')
        label.pack(side=LEFT)

        place_list = (u'santiago', u'')
        place_entry = AutocompleteEntry(place_frame)
        place_entry.set_completion_list(place_list)
        place_entry.pack()
        place_entry.focus_set()


        # L # Profesión
        occupation_frame = Frame(root)
        occupation_frame.pack()

        label = Label(occupation_frame, text=u'Profesión')
        label.pack(side=LEFT)

        occupation_list = (u'empleado', u'')
        occupation_entry = AutocompleteEntry(occupation_frame)
        occupation_entry.set_completion_list(occupation_list)
        occupation_entry.pack()
        occupation_entry.focus_set()


        # L # Street
        street_frame = Frame(root)
        street_frame.pack()

        label = Label(street_frame, text=u'Calle')
        label.pack(side=LEFT)

        street_list = (u'bandera', u'')
        street_entry = AutocompleteEntry(street_frame)
        street_entry.set_completion_list(street_list)
        street_entry.pack()
        street_entry.focus_set()


        # L # Street Number
        n_street_frame = Frame(root)
        n_street_frame.pack()

        label = Label(n_street_frame, text=u'Número de casa')
        label.pack(side=LEFT)

        n_street_entry = Entry(n_street_frame)
        n_street_entry.pack()
        n_street_entry.focus_set()        


        # Start
        root.mainloop()

if __name__ == '__main__':
        run_transcription()
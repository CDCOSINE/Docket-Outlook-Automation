# -*- coding: utf-8 -*-
"""
Created on Fri May 29 10:18:46 2020

@author: Lumenci 3
"""

from kivy.app import App
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.gridlayout import GridLayout
from kivy.uix.textinput import TextInput
from kivy.uix.screenmanager import ScreenManager, Screen
#import pickle
#import os
emm = ""
kwd = ""
class ConnectPage(GridLayout):
    def __init__(self,**kwargs):
        super().__init__(**kwargs)
        self.cols = 2
        self.add_widget(Label(text="Email address"))
        self.email = TextInput(multiline=False)
        self.add_widget(self.email)
        self.add_widget(Label(text="Enter keywords in comma sepaarted fashion"))
        self.keyword = TextInput(multiline=False)
        self.add_widget(self.keyword)
        self.join = Button(text="Submit")
        self.join.bind(on_press = self.joinbutton)
        self.add_widget(Label())
        self.add_widget(self.join)
    def joinbutton(self,instance):
        kapp.screen_manager.current = "Info"
class InfoPage(GridLayout):
    def __init__(self,**kwargs):
        super().__init__(**kwargs)
        self.cols=2
        
        self.add_widget(Label(text="Start Year"))
        self.sty = TextInput(multiline=False)
        self.add_widget(self.sty)
        
        self.add_widget(Label(text="End Year"))
        self.edy = TextInput(multiline=False)
        self.add_widget(self.edy)
        
        self.add_widget(Label(text="Start Month"))
        self.stm = TextInput(multiline=False)
        self.add_widget(self.stm)

        self.add_widget(Label(text="End Month"))
        self.edm = TextInput(multiline=False)
        self.add_widget(self.edm)       

        self.add_widget(Label(text="Start Date"))
        self.std = TextInput(multiline=False)
        self.add_widget(self.std)

        self.add_widget(Label(text="End Date"))
        self.ed = TextInput(multiline=False)
        self.add_widget(self.ed)  
        

        
        self.add_widget(Label())
        self.join = Button(text="Submit")
        #self.join.bind(on_press = self.joinbutton)
        self.add_widget(self.join)
class TestApp(App,**arg):
    def build(self):
        self.screen_manager = ScreenManager()
        
        self.connect_page = ConnectPage()
        screen = Screen(name="Connect")
        screen.add_widget(self.connect_page)
        self.screen_manager.add_widget(screen)
        
        self.info_page = InfoPage()
        screen2 = Screen(name="Info")
        screen2.add_widget(self.info_page)
        self.screen_manager.add_widget(screen2)
        
        self.screen_manager

        return self.screen_manager


if __name__ == "__main__":
    kapp = TestApp(emm,kwd)
    kapp.run()
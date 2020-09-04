#!/usr/bin/python3
# -*- coding: utf-8 -*-

from __future__ import print_function
import PySimpleGUI as sg
import Psprt_Paste_v3

sg.theme('DarkBlue12')
layout = [
          [sg.Checkbox('Сгенерировать паспорта', change_submits=True, enable_events=True, default='0',key='all')],
		  [sg.Checkbox('Обновить данные в паспортах', change_submits=True, enable_events=True, default='0',key='some')],
		  [sg.Image(r'Siba_short.png')],
          [sg.OK(), sg.Cancel()]]
window = sg.Window('Меню', layout)
event, values = window.Read()
while True:
	event, values = window.Read()
	if event == 'Cancel' or event is None:
		raise SystemExit(1)
	if values['all'] == True:
		window.Close()
		Psprt_Paste_v3.all()
		break
	if values['some'] == True:
		window.Close()
		Psprt_Paste_v3.some()
		break





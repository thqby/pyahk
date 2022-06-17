# Use AutoHotkey V2 in Python

## examples
```python
from ctypes import c_wchar_p
from pyahk import *

AhkApi.initialize()	# init ahk
AhkApi.addScript('''
A_IconHidden:=0
f4::exitapp
''')

# import ahk's global vars, or AhkApi[varname]
from pyahk import Gui, MsgBox, FileAppend, Array, Map, JSON, Hotkey

def onbutton(ob, inf):
	v = ob.Gui['Edit1'].Value
	MsgBox(v)

g = Gui()
g.AddEdit('w320 h80', 'hello python')
g.AddButton(None, 'press').OnEvent('Click', onbutton)
g.show('NA')
arr = Array(8585, 'ftgy', 85, Map(85,244,'fyg', 58))
FileAppend(JSON.stringify(arr), '*')

def onf6(key):
	MsgBox(key)

@WINFUNCTYPE(None, c_wchar_p)
def onf7(key):
	MsgBox(key)
Hotkey('F6', onf6)	# use ComProxy callback
Hotkey('F7', onf7)	# use CFuncType callback

# ahk's message pump, block until the AHK interpreter exits
AhkApi.pumpMessages()
```
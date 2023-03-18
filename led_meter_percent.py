#!/usr/bin/env python
import os, errno
import pyaudio
import numpy
from tkinter import *
from tkinter.ttk import *
import time
from tk_tools import *
from tkinter import messagebox
import ctypes
import comtypes
import sys
from ctypes import wintypes
from pynput.keyboard import Key, Controller
from time import sleep
from idlelib.tooltip import Hovertip
MMDeviceApiLib = comtypes.GUID(
    '{2FDAAFA3-7523-4F66-9957-9D5E7FE698F6}')
IID_IMMDevice = comtypes.GUID(
    '{D666063F-1587-4E43-81F1-B948E807363F}')
IID_IMMDeviceCollection = comtypes.GUID(
    '{0BD7A1BE-7A1A-44DB-8397-CC5392387B5E}')
IID_IMMDeviceEnumerator = comtypes.GUID(
    '{A95664D2-9614-4F35-A746-DE8DB63617E6}')
IID_IAudioEndpointVolume = comtypes.GUID(
    '{5CDF2C82-841E-4546-9722-0CF74078229A}')
IID_IAudioMeterInformation = comtypes.GUID('{C02216F6-8C67-4B5B-9D00-D008E73E0064}')
CLSID_MMDeviceEnumerator = comtypes.GUID(
    '{BCDE0395-E52F-467C-8E3D-C4579291692E}')
import math
import struct
# EDataFlow
eRender = 0
keyboard=Controller()
# ERole
eConsole = 0 # games, system sounds, and voice commands
eMultimedia = 1 # music, movies, narration
eCommunications = 2 # voice communications

LPCGUID = REFIID = ctypes.POINTER(comtypes.GUID)
LPFLOAT = ctypes.POINTER(ctypes.c_float)
LPDWORD = ctypes.POINTER(wintypes.DWORD)
LPUINT = ctypes.POINTER(wintypes.UINT)
LPBOOL = ctypes.POINTER(wintypes.BOOL)
PIUnknown = ctypes.POINTER(comtypes.IUnknown)
def get_resource_path(relative_path):
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
class IMMDevice(comtypes.IUnknown):
    _iid_ = IID_IMMDevice
    _methods_ = (
        comtypes.COMMETHOD([], ctypes.HRESULT, 'Activate',
            (['in'], REFIID, 'iid'),
            (['in'], wintypes.DWORD, 'dwClsCtx'),
            (['in'], LPDWORD, 'pActivationParams', None),
            (['out','retval'], ctypes.POINTER(PIUnknown), 'ppInterface')),
        comtypes.STDMETHOD(ctypes.HRESULT, 'OpenPropertyStore', []),
        comtypes.STDMETHOD(ctypes.HRESULT, 'GetId', []),
        comtypes.STDMETHOD(ctypes.HRESULT, 'GetState', []))

PIMMDevice = ctypes.POINTER(IMMDevice)

class IMMDeviceCollection(comtypes.IUnknown):
    _iid_ = IID_IMMDeviceCollection

PIMMDeviceCollection = ctypes.POINTER(IMMDeviceCollection)

class IMMDeviceEnumerator(comtypes.IUnknown):
    _iid_ = IID_IMMDeviceEnumerator
    _methods_ = (
        comtypes.COMMETHOD([], ctypes.HRESULT, 'EnumAudioEndpoints',
            (['in'], wintypes.DWORD, 'dataFlow'),
            (['in'], wintypes.DWORD, 'dwStateMask'),
            (['out','retval'], ctypes.POINTER(PIMMDeviceCollection),
             'ppDevices')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetDefaultAudioEndpoint',
            (['in'], wintypes.DWORD, 'dataFlow'),
            (['in'], wintypes.DWORD, 'role'),
            (['out','retval'], ctypes.POINTER(PIMMDevice), 'ppDevices')))
    @classmethod
    def get_default(cls, dataFlow, role):
        enumerator = comtypes.CoCreateInstance(
            CLSID_MMDeviceEnumerator, cls, comtypes.CLSCTX_INPROC_SERVER)
        return enumerator.GetDefaultAudioEndpoint(dataFlow, role)

class IAudioEndpointVolume(comtypes.IUnknown):
    _iid_ = IID_IAudioEndpointVolume
    _methods_ = (
        comtypes.STDMETHOD(ctypes.HRESULT, 'RegisterControlChangeNotify', []),
        comtypes.STDMETHOD(ctypes.HRESULT, 'UnregisterControlChangeNotify', []),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetChannelCount',
            (['out', 'retval'], LPUINT, 'pnChannelCount')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'SetMasterVolumeLevel',
            (['in'], ctypes.c_float, 'fLevelDB'),
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'SetMasterVolumeLevelScalar',
            (['in'], ctypes.c_float, 'fLevel'),
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetMasterVolumeLevel',
            (['out','retval'], LPFLOAT, 'pfLevelDB')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetMasterVolumeLevelScalar',
            (['out','retval'], LPFLOAT, 'pfLevel')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'SetChannelVolumeLevel',
            (['in'], wintypes.UINT, 'nChannel'),
            (['in'], ctypes.c_float, 'fLevelDB'),
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'SetChannelVolumeLevelScalar',
            (['in'], wintypes.UINT, 'nChannel'),
            (['in'], ctypes.c_float, 'fLevel'),
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetChannelVolumeLevel',
            (['in'], wintypes.UINT, 'nChannel'),
            (['out','retval'], LPFLOAT, 'pfLevelDB')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetChannelVolumeLevelScalar',
            (['in'], wintypes.UINT, 'nChannel'),
            (['out','retval'], LPFLOAT, 'pfLevel')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'SetMute',
            (['in'], wintypes.BOOL, 'bMute'),
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetMute',
            (['out','retval'], LPBOOL, 'pbMute')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetVolumeStepInfo',
            (['out','retval'], LPUINT, 'pnStep'),
            (['out','retval'], LPUINT, 'pnStepCount')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'VolumeStepUp',
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'VolumeStepDown',
            (['in'], LPCGUID, 'pguidEventContext', None)),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'QueryHardwareSupport',
            (['out','retval'], LPDWORD, 'pdwHardwareSupportMask')),
        comtypes.COMMETHOD([], ctypes.HRESULT, 'GetVolumeRange',
            (['out','retval'], LPFLOAT, 'pfLevelMinDB'),
            (['out','retval'], LPFLOAT, 'pfLevelMaxDB'),
            (['out','retval'], LPFLOAT, 'pfVolumeIncrementDB')))
    @classmethod
    def get_default(cls):
        endpoint = IMMDeviceEnumerator.get_default(eRender, eMultimedia)
        interface = endpoint.Activate(cls._iid_, comtypes.CLSCTX_INPROC_SERVER)
        return ctypes.cast(interface, ctypes.POINTER(cls))
class IAudioMeterInformation(comtypes.IUnknown):
    _iid_ = IID_IAudioMeterInformation
    _methods_=(comtypes.COMMETHOD([], ctypes.HRESULT, 'GetPeakValue',
            (['out','retval'], LPFLOAT, 'pfPeak')),)
    @classmethod
    def get_default(cls):
        endpoint = IMMDeviceEnumerator.get_default(eRender, eMultimedia)
        interface = endpoint.Activate(cls._iid_, comtypes.CLSCTX_INPROC_SERVER)
        return ctypes.cast(interface, ctypes.POINTER(cls))
def close():
    global appclosed
    win.destroy()
    appclosed=True
    sys.exit()
n=1
def update_limiter():
    if sb.get()=='Off':
        root.limiter=100
    else:
        root.limiter=int(sb.get())
def record_frame():
    global recording, rec_start, percent, f
    time_trunc='%.1f' % (time.time()-rec_start)
    db_str='%.2f' % percent
    f.write(time_trunc+','+db_str+'\n')
    if recording:
        win.after(100, record_frame)
    else:
        f.close()
def record():
    global recording, f, rec_start
    if not recording:
        rec.config(text='Stop')
        rec_start=time.time()
        recording=True
        f=open(os.getenv('HOMEPATH')+'\\Music\\audio_levels.csv', 'w')
        f.write('Seconds,Percent\n')
        record_frame()
    else:
        recording=False
        rec.config(text='Record')
def toggle_pre():
    global pre_fader
    pre_fader=toggle.instate(['selected'])
win=Tk()
win.title('Audio Limiter v1.0 (c) sserver')
win.grid()
recording=False
win.resizable(False, False)
win.iconbitmap(get_resource_path('snd.ico'))
tabControl = Notebook(win)
root=Frame(tabControl)
sub=Frame(tabControl)
tabControl.add(root, text ='Meter')
measure=False
percent=None
tabControl.add(sub, text ='Recording')
tabControl.pack(expand = 1, fill ="both")
gaugedb=SevenSegmentDigits(root, digits=3, digit_color='#00ff00', background='black')
gaugedb.grid(column=1, row=1)
Label(sub, text=r'Instaneous loudness (%)').grid(column=1, row=1)
recdb=SevenSegmentDigits(sub, digits=3, digit_color='#00ff00', background='black')
recdb.grid(column=1, row=2)
Hovertip(gaugedb,'Current audio level in percent')
root.limiter=80
led0 = Led(root, size=20)
led0.grid(column=2, row=13)
Hovertip(led0,'1%')
led1 = Led(root, size=20)
led1.grid(column=2, row=12)
Hovertip(led1,'10%')
led2 = Led(root, size=20)
led2.grid(column=2, row=11)
Hovertip(led2,'20%')
led3 = Led(root, size=20)
led3.grid(column=2, row=10)
Hovertip(led3,'30%')
led4 = Led(root, size=20)
led4.grid(column=2, row=9)
Hovertip(led4,'40%')
led5 = Led(root, size=20)
led5.grid(column=2, row=8)
Hovertip(led5,'50%')
led6 = Led(root, size=20)
led6.grid(column=2, row=7)
Hovertip(led6,'60%')
led7 = Led(root, size=20)
led7.grid(column=2, row=6)
Hovertip(led7,'70%')
led8 = Led(root, size=20)
led8.grid(column=2, row=5)
Hovertip(led8,'80%')
led9 = Led(root, size=20)
led9.grid(column=2, row=4)
Hovertip(led9,'90%')
led10 = Led(root, size=20)
led10.grid(column=2, row=3)
Hovertip(led10,'100%')
Label(root, text='100').grid(column=1, row=3)
Label(root, text='-').grid(column=1, row=4)
Label(root, text='80').grid(column=1, row=5)
Label(root, text='-').grid(column=1, row=6)
Label(root, text='60').grid(column=1, row=7)
Label(root, text='-').grid(column=1, row=8)
Label(root, text='40').grid(column=1, row=9)
Label(root, text='-').grid(column=1, row=10)
Label(root, text='20').grid(column=1, row=11)
Label(root, text='-').grid(column=1, row=12)
Label(root, text='Volume %').grid(column=1, row=13)
Label(root, text='Max').grid(column=3, row=0)
Label(root, text='Volume %').grid(column=1, row=0)
Label(root, text='Volume Limit').grid(column=2, row=0)
toggle=Checkbutton(root, text='Pre-Fader LEDs', command=toggle_pre)
toggle.grid(column=2, row=14)
toggle.state(['!alternate'])
list1=list(range(1, 100))
list1.append('Off')
Label(sub, text='Recording to Music\\audio_levels.csv.\nMake sure the file does not exist, or it will be overwritten.').grid(column=1, row=3)
rec=Button(sub, text='Record', command=record)
rec.grid(column=1, row=4)
sb=Spinbox(root, width=8, state='readonly', values=list1, command=update_limiter)
sb.grid(column=2, row=1)
sb.set('Off')
pre_fader=False
maxdb_display=SevenSegmentDigits(root, digits=3, digit_color='#00ff00', background='black')
maxdb_display.grid(column=3, row=1)
Hovertip(maxdb_display,'Max volume in % while app was running')
appclosed=False
Hovertip(sb,'Maximum loudness before volume is reduced automatically')
Hovertip(toggle,'Indicate loudness pre-volume control on the LEDs. CSV recording, volume limiting, and 7-segment displays will be unaffected.')
max_percent=0
def listen(old=0, error_count=0, min_decibel=100):
    global max_percent, percent, pre_fader
    global appclosed
    ev = IAudioEndpointVolume.get_default()
    listener=IAudioMeterInformation.get_default()
    if not appclosed:
        x=listener.GetPeakValue()*100
        percent=round(x)*(round(ev.GetMasterVolumeLevelScalar()*100)/100)*(1-ev.GetMute())
        percent_pre=round(x)*(1-ev.GetMute())
        if round(percent*10)/10>=100:
            gaugedb.set_value('100')
            recdb.set_value('100')
        else:
            gaugedb.set_value(str(round(percent*10)/10))
            recdb.set_value(str(round(percent*10)/10))
        if float(percent)>float(max_percent):
            max_percent=percent
        if round(max_percent*10)/10>=100:
            maxdb_display.set_value('100')
        else:
            maxdb_display.set_value(str(round(max_percent*10)/10))
        if pre_fader:
            percent_led=percent_pre
        else:
            percent_led=percent
        if int(percent_led)>0:
            led0.to_green(on=True)
        else:
            led0.to_grey(on=True)
        for i in range(1, 11):
            if int(percent_led)>=(10*i):
                if i>=9:
                    exec("led"+str(i)+".to_red(on=True)")
                elif i>=7:
                    exec("led"+str(i)+".to_yellow(on=True)")
                else:
                    exec("led"+str(i)+".to_green(on=True)")
            else:
                exec("led"+str(i)+".to_grey(on=True)")
        if percent>root.limiter:
            keyboard.tap(Key.media_volume_down)
        win.after(50, listen)
win.protocol('WM_DELETE_WINDOW', close)
if __name__ == '__main__':
    listen()
    if appclosed:
        win.destroy()
    try:
        win.mainloop()
    except TclError:
        pass

import pywintypes
import win32api
import win32con
import psutil
import time

def get_refresh_rates_for_native_resolution():
    device_name = win32api.EnumDisplayDevices(None, 0).DeviceName
    max_resolution = (0, 0)
    resolution_refresh_rates = []
    i = 0
    while True:
        try:
            display_settings = win32api.EnumDisplaySettings(device_name, i)
            if (display_settings.PelsWidth, display_settings.PelsHeight) > max_resolution:
                max_resolution = (display_settings.PelsWidth, display_settings.PelsHeight)
                resolution_refresh_rates = [display_settings.DisplayFrequency]
            elif (display_settings.PelsWidth, display_settings.PelsHeight) == max_resolution:
                resolution_refresh_rates.append(display_settings.DisplayFrequency)
            i += 1
        except win32api.error:
            break
    return resolution_refresh_rates


def change_refresh_rate(is_plugged_in: bool):
    refresh_rates = get_refresh_rates_for_native_resolution()
    #print(refresh_rates)
    devmode = pywintypes.DEVMODEType()
    if is_plugged_in:
        print('set to 240hz')
        devmode.DisplayFrequency = (max(refresh_rates))
    elif not is_plugged_in:
        print('set to 60hz')
        devmode.DisplayFrequency = (min(refresh_rates))

    devmode.Fields = win32con.DM_DISPLAYFREQUENCY

    win32api.ChangeDisplaySettings(devmode, 0)


plugged_in_state = psutil.sensors_battery().power_plugged
change_refresh_rate(plugged_in_state)
while True:
    loop_state = psutil.sensors_battery().power_plugged
    if plugged_in_state == loop_state:
        print("No state change")
        time.sleep(5)
    elif plugged_in_state != loop_state:
        print("State Changed")
        plugged_in_state = loop_state
        change_refresh_rate(loop_state)
        time.sleep(5)

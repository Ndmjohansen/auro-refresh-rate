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


def get_refresh_rate():
    device = win32api.EnumDisplayDevices()
    settings = win32api.EnumDisplaySettings(device.DeviceName, -1)
    return settings.DisplayFrequency


def flip_refresh_rate(current_hz, new_hz):
    while current_hz != new_hz:
        # print("flip refresh loop")
        devmode = pywintypes.DEVMODEType()
        devmode.DisplayFrequency = new_hz
        devmode.Fields = win32con.DM_DISPLAYFREQUENCY
        win32api.ChangeDisplaySettings(devmode, 0)
        # Wait for device to flip refresh rate
        time.sleep(3)
        current_hz = get_refresh_rate()
    # print('Done flipping refresh rate')


# Global variables
plugged_in_state = psutil.sensors_battery().power_plugged
refresh_rates = get_refresh_rates_for_native_resolution()
min_refresh = min(refresh_rates)
max_refresh = max(refresh_rates)
# set correct refresh rate on startup
if plugged_in_state:
    flip_refresh_rate(get_refresh_rate(), max_refresh)
elif not plugged_in_state:
    flip_refresh_rate(get_refresh_rate(), min_refresh)

# Main loop
while True:
    loop_state = psutil.sensors_battery().power_plugged
    if plugged_in_state == loop_state:
        # print("No state change")
        time.sleep(5)
    elif plugged_in_state != loop_state:
        # print("State Changed")
        plugged_in_state = loop_state
        if plugged_in_state:
            flip_refresh_rate(get_refresh_rate(), max_refresh)
        elif not plugged_in_state:
            flip_refresh_rate(get_refresh_rate(), min_refresh)
        time.sleep(5)

import os
import time
import logging
import shutil
from os import walk
from os.path import exists

import py7zr
import requests
import usb.core
import usb.util
from dotenv import load_dotenv
from usb.backend import libusb1


load_dotenv()

# declare constants
# 1. product ID
# Razer Viper V2 Pro (Wired) 	1532:00A5
# Razer Viper V2 Pro (Wireless) 	1532:00A6
# see README.md for instruction to find the device ID for your mouse
WIRELESS_RECEIVER = 0x00A6
WIRELESS_WIRED = 0x00A5
# 2. transaction_id.id
# 0x3f for Razer Viper Wireless
# see README.md for instruction to find the correct transaction_id.id for your mouse
TRAN_ID = b"\x3f"

HA_IP = os.getenv('ha_ip') 
HA_PORT = os.getenv('ha_port') 
HA_TOKEN = os.getenv('ha_token') 
HA_URL = f"http://{HA_IP}:{HA_PORT}/api/states/binary_sensor.mouse_low"


def get_libusb(is_latest=False):
    """
    Downloads libusb dll file and moves it to System32 and SysWOW64
    Args:
        is_latest (bool): Grab latest if True, else grab default version
    Returns:

    """
    file_exists = exists("C:/Windows/System32/libusb-1.0.dll") and exists("C:/Windows/SysWOW64/libusb-1.0.dll")
    if file_exists:
        return

    # url = "https://github.com/libusb/libusb/releases/download/v1.0.27/libusb-1.0.27.7z"
    version = "v1.0.27"
    latest_vs = "VS2020"
    if is_latest:
        libusb_latest_url = "https://github.com/libusb/libusb/releases/latest"
        libusb_version_response = requests.get(libusb_latest_url, timeout=10)
        version = libusb_version_response.url.split("/").pop()

        latest_vs = list(reversed(list(filter(lambda x: x.startswith("VS"), list(walk("tmp"))[0][1]))))[0]

    libusb_url = f"https://github.com/libusb/libusb/releases/download/{version}/libusb-{version[1:]}.7z"
    libusb_response = requests.get(libusb_url, stream=True, timeout=10)

    if libusb_response.status_code == 200:
        with open("libusb.7z", 'wb') as out:
            out.write(libusb_response.content)
    else:
        logging.info("Request failed: %d", libusb_response.status_code)

    with py7zr.SevenZipFile("libusb.7z", 'r') as archive:
        archive.extractall(path="tmp")

    shutil.move(f"tmp/{latest_vs}/MS64/dll/libusb-1.0.dll", "C:/Window/System32/libusb-1.0.dll")
    shutil.move(f"tmp/{latest_vs}/MS32/dll/libusb-1.0.dll", "C:/Windows/SysWOW64/libusb-1.0.dll")

    shutil.rmtree('tmp')


def get_mouse():
    """
    Function that checks whether the mouse is plugged in or not
    :return: [mouse, wireless]: a list that stores (1) a Device object that represents the mouse; and
    (2) a boolean for stating if the mouse is in wireless state (True) or wired state (False)
    """
    # declare backend: libusb1.0
    backend = libusb1.get_backend()
    # find the mouse by PyUSB
    mouse = usb.core.find(idVendor=0x1532, idProduct=WIRELESS_RECEIVER, backend=backend)
    # if the receiver is not found, mouse would be None
    if not mouse:
        # try finding the wired mouse
        mouse = usb.core.find(idVendor=0x1532, idProduct=WIRELESS_WIRED, backend=backend)
        # still not found, then the mouse is not plugged in, raise error
        if not mouse:
            raise RuntimeError(f"The specified mouse (PID:{WIRELESS_RECEIVER} or {WIRELESS_WIRED}) cannot be found.")

    return mouse


def battery_msg():
    """
    Function that creates and returns the message to be sent to the device
    Args:
        
    Returns:
        msg: the message to be sent to the mouse for getting the battery level
    """
    # adapted from https://github.com/rsmith-nl/scripts/blob/main/set-ornata-chroma-rgb.py
    # the first 8 bytes in order from left to right
    # status + transaction_id.id + remaining packets (\x00\x00) + protocol_type + command_class + command_id + data_size
    msg = b"\x00" + TRAN_ID + b"\x00\x00\x00\x02\x07\x80"
    crc = 0
    for i in msg[2:]:
        crc ^= i
    # the next 80 bytes would be storing the data to be sent, but for getting the battery no data is sent
    msg += bytes(80)
    # the last 2 bytes would be the crc and a zero byte
    msg += bytes([crc, 0])
    return msg


def get_battery():
    """
    Function for getting the battery level of a Razer Viper Wireless, or other device if adapted
    Args:
        
    Returns:
        a float with the battery level as a percentage
    """
    # find the mouse and the state, see get_mouse() for detail
    mouse = get_mouse()
    # the message to be sent to the mouse, see battery_msg() for detail
    msg = battery_msg()
    logging.info("Message sent to the mouse: %s", list(msg))
    # needed by PyUSB
    # if Linux, need to detach kernel driver
    mouse.set_configuration()
    usb.util.claim_interface(mouse, 0)
    # send request (battery), see razer_send_control_msg in razercommon.c in OpenRazer driver for detail
    req = mouse.ctrl_transfer(bmRequestType=0x21, bRequest=0x09, wValue=0x300, data_or_wLength=msg,
                              wIndex=0x00)
    # needed by PyUSB
    usb.util.dispose_resources(mouse)
    # will wait before getting response regardless of wireless or not
    time.sleep(0.3305)
    # receive response
    result = mouse.ctrl_transfer(bmRequestType=0xa1, bRequest=0x01, wValue=0x300, data_or_wLength=90, wIndex=0x00)
    usb.util.dispose_resources(mouse)
    usb.util.release_interface(mouse, 0)
    logging.info("Message received from the mouse: %s", list(result))
    # the raw battery level is in 0 - 255, scale it to 100 for human, correct to 2 decimal places
    return result[9] / 255 * 100


def update_ha(battery):
    """
    Updates the binary sensor on HA
    Args:
        battery (float): Grab latest if True, else grab default version
    Returns:

    """
    current_state = "off"
    if battery < 42:
        current_state = "on"

    response = requests.get(
        HA_URL,
        headers={
            "Authorization": f"Bearer {HA_TOKEN}",
            },
        timeout=10,
        )

    state_match = False
    if response.status_code == 200:
        state_on_ha = response.json().get("state")
        if state_on_ha and state_on_ha == current_state:
            state_match = True

    if not state_match:
        response = requests.post(
            HA_URL,
            headers={
                "Authorization": f"Bearer {HA_TOKEN}",
                "content-type": "application/json",
                },
            json={"state": current_state, "attributes": {"friendly_name": "MouseLow"}},
            timeout=10,
            )


if __name__ == "__main__":
    # get_libusb()
    battery = get_battery()
    logging.info("Battery level obtained: %s", battery)
    update_ha(battery)

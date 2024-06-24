#!/usr/bin/env python3

"""
Handle automatic connection with MPP Solar inverter compatible device (VEVOR)
This will output 2 dbus services, one for Inverter data another one for control
via VRM of the features.
"""
VERSION = 'v0.2' 

from gi.repository import GLib
import platform
import argparse
import logging
import sys
import os
import subprocess as sp
import json
from enum import Enum
import datetime
import dbus
import dbus.service

logging.basicConfig(level=logging.WARNING)

# our own packages
sys.path.insert(1, os.path.join(os.path.dirname(__file__), 'velib_python'))
from vedbus import VeDbusService, VeDbusItemExport, VeDbusItemImport

# Workarounds for some inverter specific problem I saw
INVERTER_OFF_ASSUME_BYPASS = True
GUESS_AC_CHARGING = True

# Should we import and call manually, to use our version
USE_SYSTEM_MPPSOLAR = False
if USE_SYSTEM_MPPSOLAR:
    try:
        import mppsolar
    except:
        USE_SYSTEM_MPPSOLAR = FALSE
if not USE_SYSTEM_MPPSOLAR:
    sys.path.insert(1, os.path.join(os.path.dirname(__file__), 'mpp-solar'))
    import mppsolar

# Inverter commands to read from the serial
def runInverterCommands(commands, protocol="PI18"):
    global args
    global mainloop
    if USE_SYSTEM_MPPSOLAR:
        output = [sp.getoutput("mpp-solar -b {} -P {} -p {} -o json -c {}".format(args.baudrate, protocol, args.serial, c)).split('\n')[0] for c in commands]
        parsed = [json.loads(o) for o in output]
    else:
        dev = mppsolar.helpers.get_device_class("mppsolar")(port=args.serial, protocol=protocol, baud=args.baudrate)
        results = [dev.run_command(command=c) for c in commands]
        parsed = [mppsolar.outputs.to_json(r, False, None, None) for r in results]           
    return parsed

def setOutputSource(source, protocol="PI18"):
    #POP<NN>: Setting device output source priority
    #    NN = 00 for utility first, 01 for solar first, 02 for SBU priority
    #   For PI18, Output POP0 [0: Solar-Utility-Batter],  POP1 [1: Solar-Battery-Utility]
    return runInverterCommands(['POP{:02d}'.format(source)])

def setChargerPriority(priority, protocol="PI18"):
    #PCP<NN>: Setting device charger priority
    #  For KS: 00 for utility first, 01 for solar first, 02 for solar and utility, 03 for only solar charging
    #  For MKS: 00 for utility first, 01 for solar first, 03 for only solar charging
    #   For PI18, 0: Solar first, 1: Solar and Utility, 2: Only solar
    return runInverterCommands(['PCP{:02d}'.format(priority)])

def setMaxChargingVoltage(voltage, protocol="PI18"):
    #MCHGV : Setting bulk and float voltage
    # For PI18 : MCHGV552,540 will set Bulk - CV voltage [480~584] in 0.1V xxx, Float voltage [480~584] in 0.1V
    if protocol == "PI18":
        return runInverterCommands(['MCHGV{:d},{:d}'.format(int(voltage*10), int(voltage*10))], protocol)
    else:
        return True

def setMaxChargingCurrent(current, protocol="PI18"):
    #MNCHGC<mnnn><cr>: Setting max charging current (More than 100A)
    #  Setting value can be gain by QMCHGCR command.
    #  nnn is max charging current, m is parallel number.
    roundedCurrent = round(current / 10) * 10
    if protocol == "PI18":
        return runInverterCommands(['MCHGC0{:04d}'.format(roundedCurrent)], protocol)
    else:
        return runInverterCommands(['MNCHGC0{:04d}'.format(roundedCurrent)], protocol)

def setMaxUtilityChargingCurrent(current, protocol="PI18"):
    #MUCHGC<nnn><cr>: Setting utility max charging current
    #  Setting value can be gain by QMCHGCR command.
    #  nnn is max charging current, m is parallel number.
    roundedCurrent = max(2, round(current / 10) * 10)
    if protocol == "PI18":
        return runInverterCommands(['MUCHGC0{:04d}'.format(roundedCurrent)], protocol)
    else:
        return runInverterCommands(['MUCHGC{:03d}'.format(current)])

def isNaN(num):
    return num != num


# Allow to have multiple DBUS connections
class SystemBus(dbus.bus.BusConnection):
    def __new__(cls):
        return dbus.bus.BusConnection.__new__(cls, dbus.bus.BusConnection.TYPE_SYSTEM) 
class SessionBus(dbus.bus.BusConnection):
    def __new__(cls):
        return dbus.bus.BusConnection.__new__(cls, dbus.bus.BusConnection.TYPE_SESSION)
def dbusconnection():
    return SessionBus() if 'DBUS_SESSION_BUS_ADDRESS' in os.environ else SystemBus()

# Our MPP solar service that connects to 2 dbus services (multi & vebus)
class DbusMppSolarService(object):
    def __init__(self, tty, deviceinstance, productname='MPPSolar', connection='MPPSolar interface', json_file_path='/data/etc/dbus-mppsolar/config.json'):
        self._tty = tty
        self._queued_updates = []
        
        # Get the name from config file if available
        if os.path.exists(json_file_path):
            with open(json_file_path, 'r') as json_file:
                config = json.load(json_file)
            if tty in config:
                productname_value = config[self._tty].get('productname', None)
                if productname_value is not None:
                    productname = productname_value
                    logging.warning("Product named from config : {}".format(productname_value))

        self._invProtocol = 'PI18'
        logging.warning(f"Protocol set to PI18.")
        
        # Get inverter data based on protocol
        try:
            self._invData = runInverterCommands(['ID','VFW'], self._invProtocol)
        except:
            logging.warning(f"Error getting datas from inverter on {tty} ({self._invProtocol})")

        logging.warning(f"Connected to inverter on {tty} ({self._invProtocol}), setting up dbus with /DeviceInstance = {deviceinstance}")
        
        # Create the services
        self._dbusinverter = VeDbusService(f'com.victronenergy.inverter.mppsolar-inverter.{tty}', dbusconnection())
        # self._dbusvebus = VeDbusService(f'com.victronenergy.vebus.mppsolar.{tty}', dbusconnection())
        self._dbusmppt = VeDbusService(f'com.victronenergy.solarcharger.mppsolar-charger.{tty}', dbusconnection())

        # Set up default paths
        self.setupInverterDefaultPaths(self._dbusinverter, connection, deviceinstance, f"Inverter {productname}")
        # self.setupInverterDefaultPaths(self._dbusvebus, connection, deviceinstance, f"Vebus {productname}")
        self.setupChargerDefaultPaths(self._dbusmppt, connection, deviceinstance, f"Charger {productname}")

        # self._system = VeDbusItemImport(dbusconnection(), f'com.victronenergy.system', '/Connected')

        # Create paths for inverter
        self._dbusinverter.add_path('/Dc/0/Voltage', 0)
        self._dbusinverter.add_path('/Ac/Out/L1/V', 0)
        self._dbusinverter.add_path('/Ac/Out/L1/I', 0)
        self._dbusinverter.add_path('/Ac/Out/L1/P', 0)
        self._dbusinverter.add_path('/Mode', 0)                     #<- Switch position: 2=Inverter on; 4=Off; 5=Low Power/ECO
        self._dbusinverter.add_path('/State', 0)                    #<- 0=Off; 1=Low Power; 2=Fault; 9=Inverting
        self._dbusinverter.add_path('/Temperature', 123)
           
        logging.info(f"Paths for Inverter created.")

        # Create paths for charger
        # general data
        self._dbusmppt.add_path('/NrOfTrackers', 1)
        self._dbusmppt.add_path('/Pv/V', 0)
        # self._dbusmppt.add_path('/Pv/0/V', 0)
        # self._dbusmppt.add_path('/Pv/0/P', 0)
        self._dbusmppt.add_path('/Yield/Power', 0)
        self._dbusmppt.add_path('/DC/0/Temperature', 123)
        self._dbusmppt.add_path('/Dc/0/Voltage', 0)
        self._dbusmppt.add_path('/Dc/0/Current', 0)

        # external control
        self._dbusmppt.add_path('/Link/NetworkMode', 1) # <- Bitmask
                        # 0x1 = External control
                        # 0x4 = External voltage/current control
                        # 0x8 = Controled by BMS (causes Error #67, BMS lost, if external control is interrupted).
        self._dbusmppt.add_path('/Link/BatteryCurrent', 0)
        self._dbusmppt.add_path('/Link/ChargeCurrent', 0)
        self._dbusmppt.add_path('/Link/ChargeVoltage', 0)
        self._dbusmppt.add_path('/Link/NetworkStatus', 4) # <- Bitmask
                        # 0x01 = Slave
                        # 0x02 = Master
                        # 0x04 = Standalone
                        # 0x20 = Using I-sense (/Link/BatteryCurrent)
                        # 0x40 = Using T-sense (/Link/TemperatureSense)
                        # 0x80 = Using V-sense (/Link/VoltageSense)
        self._dbusmppt.add_path('/Link/TemperatureSense', 0)
        self._dbusmppt.add_path('/Link/TemperatureSenseActive', 0)
        self._dbusmppt.add_path('/Link/VoltageSense', 0)
        self._dbusmppt.add_path('/Link/VoltageSenseActive', 0)
        # settings
        self._dbusmppt.add_path('/Settings/BmsPresent', None)
        self._dbusmppt.add_path('/Settings/ChargeCurrentLimit', 80)
        # other paths
        self._dbusmppt.add_path('/Yield/User', 0)
        self._dbusmppt.add_path('/Yield/System', 0)
        self._dbusmppt.add_path('/ErrorCode', 0)
        self._dbusmppt.add_path('/State', 0)
        self._dbusmppt.add_path('/Mode', 0)
        self._dbusmppt.add_path('/MppOperationMode', 0)
        self._dbusmppt.add_path('/Relay/0/State', None)
        # history
        # self._dbusmppt.add_path('/History/Overall/DaysAvailable', 0)
        # self._dbusmppt.add_path('/History/Overall/MaxPvVoltage', 0)
        # self._dbusmppt.add_path('/History/Overall/MaxBatteryVoltage', 0)
        # self._dbusmppt.add_path('/History/Overall/MinBatteryVoltage', 0)

        logging.info(f"Paths for 'solarcharger' created.")

        # Create paths for 'vebus'
        # self._dbusvebus.add_path('/Ac/ActiveIn/L1/F', 0)
        # self._dbusvebus.add_path('/Ac/ActiveIn/L1/I', 0)
        # self._dbusvebus.add_path('/Ac/ActiveIn/L1/V', 0)
        # self._dbusvebus.add_path('/Ac/ActiveIn/L1/P', 0)
        # self._dbusvebus.add_path('/Ac/ActiveIn/L1/S', 0)
        # self._dbusvebus.add_path('/Ac/ActiveIn/P', 0)
        # self._dbusvebus.add_path('/Ac/ActiveIn/S', 0)
        # self._dbusvebus.add_path('/Ac/ActiveIn/ActiveInput', 0)

        # self._dbusvebus.add_path('/Ac/Out/L1/V', 0)
        # self._dbusvebus.add_path('/Ac/Out/L1/I', 0)
        # self._dbusvebus.add_path('/Ac/Out/L1/P', 0)
        # self._dbusvebus.add_path('/Ac/Out/L1/S', 0)
        # self._dbusvebus.add_path('/Ac/Out/L1/F', 0)

        # self._dbusvebus.add_path('/Ac/NumberOfPhases', 1)
        # self._dbusvebus.add_path('/Dc/0/Voltage', 0)
        # self._dbusvebus.add_path('/Dc/0/Current', 0)
        # self._dbusvebus.add_path('/Pv/0/V',0)
        # self._dbusvebus.add_path('/Pv/V',0)
        # self._dbusvebus.add_path('/Pv/0/P',0)
        # self._dbusvebus.add_path('/Yield/Power',0)
        # self._dbusvebus.add_path('/MppOperationMode',0)

        # self._dbusvebus.add_path('/Ac/In/1/CurrentLimit', 20, writeable=True, onchangecallback=self._change)
        # self._dbusvebus.add_path('/Ac/In/1/CurrentLimitIsAdjustable', 1)
        # self._dbusvebus.add_path('/Settings/SystemSetup/AcInput1', 1)
        # self._dbusvebus.add_path('/Ac/In/1/Type', 1) #0=Unused;1=Grid;2=Genset;3=Shore
        
        # self._dbusvebus.add_path('/Mode', 0, writeable=True, onchangecallback=self._change)
        # self._dbusvebus.add_path('/ModeIsAdjustable', 1)
        # self._dbusvebus.add_path('/State', 0)
        # self._dbusvebus.add_path('/Ac/In/1/L1/V', 0, writeable=False, onchangecallback=self._change)

        GLib.timeout_add(10000 if USE_SYSTEM_MPPSOLAR else 10000, self._update)
    
    def setupInverterDefaultPaths(self, service, connection, deviceinstance, productname):
        # Create the management objects, as specified in the ccgx dbus-api document
        service.add_path('/Mgmt/ProcessName', __file__)
        service.add_path('/Mgmt/ProcessVersion', 'version f{VERSION}, and running on Python ' + platform.python_version())
        service.add_path('/Mgmt/Connection', connection)

        # Create the mandatory objects
        service.add_path('/DeviceInstance', deviceinstance)
        service.add_path('/ProductId', None)
        service.add_path('/ProductName', productname)
        service.add_path('/FirmwareVersion', None)
        service.add_path('/HardwareVersion', None)
        service.add_path('/Connected', 1)

        # Create the paths for modifying the system manually
        service.add_path('/Settings/Reset', None, writeable=True, onchangecallback=self._change)
        service.add_path('/Settings/Charger', None, writeable=True, onchangecallback=self._change)
        service.add_path('/Settings/Output', None, writeable=True, onchangecallback=self._change)

    def setupChargerDefaultPaths(self, service, connection, deviceinstance, productname):
        # Create the management objects, as specified in the ccgx dbus-api document
        service.add_path('/Mgmt/ProcessName', __file__)
        service.add_path('/Mgmt/ProcessVersion', 'version f{VERSION}, and running on Python ' + platform.python_version())
        service.add_path('/Mgmt/Connection', connection)

        # Create the mandatory objects
        service.add_path('/DeviceInstance', deviceinstance)
        service.add_path('/ProductId', None)
        service.add_path('/ProductName', productname)
        service.add_path('/FirmwareVersion', None)
        service.add_path('/HardwareVersion', None)
        service.add_path('/Connected', 1)

        # Create the paths for modifying the system manually
        service.add_path('/Settings/Reset', None, writeable=True, onchangecallback=self._change)
        service.add_path('/Settings/Charger', None, writeable=True, onchangecallback=self._change)
        service.add_path('/Settings/Output', None, writeable=True, onchangecallback=self._change)

    def _updateInternal(self):
        # Store in the paths all values that were updated from _handleChangedValue
        with self._dbusinverter as i, self._dbusmppt as m:# self._dbusvebus as v:
            for path, value, in self._queued_updates:
                i[path] = value
                m[path] = value
            self._queued_updates = []

    def _update(self):
        global mainloop
        logging.info("{} updating".format(datetime.datetime.now().time()))
        try: 
            return self._update_PI18()
        except:
            logging.exception('Error in update loop', exc_info=True)
            # mainloop.quit()
            self._updateInternal()
            return True

    def _change(self, path, value):
        global mainloop
        logging.warning("updated %s to %s" % (path, value))
        if path == '/Settings/Reset':
            logging.info("Restarting!")
            mainloop.quit()
            exit
        try: 
            return self._change_PI18(path, value)
        except:
            logging.exception('Error in change loop', exc_info=True)
            mainloop.quit()
            return False

    def _update_PI18(self):
       # raw = runInverterCommands(['GS','MOD','FWS'])
        try:
            raw = runInverterCommands(['ET','GS','MOD','PIRI'], "PI18")
            # logging.warning(raw)
        except:
            logging.warning("Error in update PI18 loop.", exc_info=True)
            self._updateInternal()
            return True
        
    # data, mode, warnings = raw
        generated, data, mode, rated = raw
        with self._dbusinverter as i, self._dbusmppt as m:           # self._dbusvebus as v,
            # 0=Off;1=Low Power;2=Fault;9=Inverting
            invMode = mode.get('working_mode', i['/State'])
            if invMode == 'Battery mode':
                i['/State'] = 9 # Inverting
            elif invMode == 'Fault mode':
                i['/State'] = 2 # Fault mode
            else:
                i['/State'] = 0 # OFF

            # Normal operation, read data
            i['/Dc/0/Voltage'] = data.get('battery_voltage', i['/Dc/0/Voltage'])
            m['/Dc/0/Voltage'] = data.get('battery_voltage', m['/Dc/0/Voltage'])

            # i['/Dc/0/Current'] = data.get('battery_charging_current', 0) - data.get('battery_discharge_current', 0)

            i['/Ac/Out/L1/V'] = data.get('ac_output_voltage', i['/Ac/Out/L1/V'])
            i['/Ac/Out/L1/P'] = data.get('ac_output_active_power', i['/Ac/Out/L1/P'])
            if i['/Ac/Out/L1/V'] != 0 & i['/Ac/Out/L1/P'] != 0:
                output_current = i['/Ac/Out/L1/P'] / i['/Ac/Out/L1/V']
                i['/Ac/Out/L1/I'] = output_current

            # Solar charger
            if data.get('pv1_input_power', 0) > 0:
                m['/State'] = 3
            else:
                m['/State'] = 0
            # m['/Pv/0/V'] = data.get('pv1_input_voltage', m['/Pv/0/V'])
            m['/Pv/V'] = data.get('pv1_input_voltage', m['/Pv/V'])
            # m['/Pv/0/P'] = data.get('pv1_input_power', m['/Pv/0/P'])
            m['/Yield/Power'] = data.get('pv1_input_power', m['/Yield/Power'])
            m['/Yield/User'] = generated.get('total_generated_energy', m['/Yield/User']) / 1000
            m['/Yield/System'] = generated.get('total_generated_energy', m['/Yield/System']) / 1000
            m['/MppOperationMode'] = 2 if (data.get('pv1_input_power') != None and data.get('pv1_input_power') > 0) else 0
            #m['/Link/ChargeCurrent'] =  rated.get('max_charging_current',  m['/Link/ChargeCurrent']) # <- Maximum charge current. Must be written every 60 seconds. Used by GX device if there is a BMS or user limit.
            #m['/Link/ChargeVoltage'] =  rated.get('battery_bulk_voltage',  m['/Link/ChargeVoltage']) # <- Charge voltage. Must be written every 60 seconds. Used by GX device to communicate BMS charge voltages.
            
            # # Misc
            i['/Temperature'] = data.get('inverter_heat_sink_temperature', i['/Temperature'])
            m['/DC/0/Temperature'] = data.get('mppt1_charger_temperature', m['/DC/0/Temperature'])

            # # Execute updates of previously updated values
            self._updateInternal()
        return True

    def _change_PI18(self, path, value):
        # Link
        if path == '/Link':
            logging.warning("{} : {}".format(path, value))

        if path == '/Link/ChargeCurrent':
            logging.warning("/Link/ChargeCurrent : {}".format(value))

        if path == '/Link/ChargeCurrent':
            logging.warning("/Link/ChargeCurrent : {}".format(value))

        # Mode settings
        if path == '/Mode': # 1=Charger Only;2=Inverter Only;3=On;4=Off(?)
            if value == 1:
                #logging.warning("setting mode to 'Charger Only'(Charger=Util & Output=Util->solar) ({},{})".format(setChargerPriority(0), setOutputSource(0)))
                logging.warning("setting mode to 'Charger Only'(Charger=Util) ({})".format(setChargerPriority(1), setOutputSource(1)))
            elif value == 2:
                logging.warning("setting mode to 'Inverter Only'(Charger=Solar & Output=SBU) ({},{})".format(setChargerPriority(0), setOutputSource(2)))
            elif value == 3:
                logging.warning("setting mode to 'ON=Charge+Invert'(Charger=Util & Output=SBU) ({},{})".format(setChargerPriority(1), setOutputSource(2)))
            elif value == 4:
                #logging.warning("setting mode to 'OFF'(Charger=Solar & Output=Util->solar) ({},{})".format(setChargerPriority(3), setOutputSource(0)))
                logging.warning("setting mode to 'OFF'(Charger=Solar) ({})".format(setChargerPriority(3), setOutputSource(2)))
            else:
                logging.warning("setting mode not understood ({})".format(value))
            self._queued_updates.append((path, value))        
        return True # accept the change

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--baudrate","-b", default=2400, type=int)
    parser.add_argument("--serial","-s", required=True, type=str)
    global args
    args = parser.parse_args()

    from dbus.mainloop.glib import DBusGMainLoop
    # Have a mainloop, so we can send/receive asynchronous calls to and from dbus
    DBusGMainLoop(set_as_default=True)

    mppservice = DbusMppSolarService(tty=args.serial.strip("/dev/"), deviceinstance=0)
    logging.warning('Created service & connected to dbus, switching over to GLib.MainLoop() (= event based)')

    global mainloop
    mainloop = GLib.MainLoop()
    mainloop.run()
    

if __name__ == "__main__":
    main()

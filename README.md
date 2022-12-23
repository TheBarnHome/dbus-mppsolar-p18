# dbus-vevor-inverter
DBus VenusOS driver for Vevor inverter

# INSTRUCTIONS

- Clone GIT & submodules:
  -- git clone --recurse-submodules https://github.com/DarkZeros/dbus-vevor-inverter /data/etc/dbus-vevor-inverter
  -- Need velib_python for execution of the service
  -- Need mpp-solar to communicate with Inverter

- Symlink & install service
   ln -s /data/etc/dbus-vevor-inverter/service-templates /opt/victronenergy/service-templates/dbus-vevor-inverter

- Add service to  /etc/venus/serial-starter.conf and modify the default list:
...
service vevor           dbus-vevor-inverter
alias   default         gps:vedirect:vevor
...

- Add udev rule to run the service only if you connect your VEVOR inverter:

/etc/udev/rules.d/serial-starter.rules

ACTION=="add", ENV{ID_BUS}=="usb", ATTRS{idVendor}=="067b", ATTRS{serial}=="ELARb11A920",          ENV{VE_SERVICE}="vevor"
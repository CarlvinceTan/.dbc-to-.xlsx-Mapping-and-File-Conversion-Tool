.dbc to .xlsx Mapping and File Conversion Tool
============================================
**Written by Carlvince Tan for MUR Documentation Tools**

## Motivation
- This tool is mainly to enable ease of access and maintainability for electrical engineers (specifically pin mapping); however, for software engineers who require access to further information about CAN messages, referring to the .dbc file will still be most advisable.
- The demo Vector CANdb++ software isn't able to create new attributes/categories to specify pins, adding these extra details in the comments will also result in permanent adjustments to the .dbc, unable to be uncommented.

## Features
- Auto-renaming
- Automatic cell-width adjustment and formatting
- Dependecies are pre-included

## How to Use
Go to .dbcConverter.py file. (Ignore IDE warnings)

Place only one .dbc file into the Input folder, adjust/add/remove CAN connections listed in the pin_config.txt file in the following format (excluding square brackets):
```Bash
[Device1] -> [Device2]
High: [Pin1] -> [Pin2], Low: [Pin3] -> [Pin4]
[newline]
```
*N.B. It is whitespace sensitive*

Then run the following in your terminal according to your operating system:

Windows:
- Run `./win.sh`

Mac/Linux:
- Run `./mac_linux.sh`

The output should appear in the Output folder

## Resolving issues
If your terminal requests permission to execute shell script:

Windows:
- Run `icacls bash_script.sh /grant Users:(RX)`

Mac/Linux:
- Run `chmod +x bash_script.sh`

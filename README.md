# RunGameIso

This app uses a config file to mount an ISO in daemon tools, run a program, and optional unmount it after. It can be used to create shortcuts to start games or other programs.

I generally use this tool by creating a shortcut to the executable, assigning it an icon, and changing the shortcut to provide the INI file path parameter.

### Code

This project was coded with Visual Basic 6 on a Windows 98 SE machine.

### How to run

Run the executable and pass (without any qoutes) the path to the config file.

```terminal
VirtualDaemonLauncher.exe C:\IsoConfigs\MyGame.ini
```

### Config file

The config file used is an INI file with a simple format. Most items can be omitted from the file, as defaults will be provided.

```ini
[DEFAULT]
game_name=Arthur's Teacher Trouble
daemon_executable=c:\program files\d-tools\daemon.exe
device_number=0
image=F:\iso\arthur's teacher trouble\arthur.cue
wait_seconds_before_program=10
program=H:\arthur.exe
safedisc=off
securom=off
laserlock=off
rmps=off
unmount=true
```

| Entry                         | Description                                                                                                       |
|-------------------------------|-------------------------------------------------------------------------------------------------------------------|
| `game_name`                   | A title used to display in a simple popup dialog while the CD-ROM image is mounted and before the program is run. |
| `daemon_executable`           | The path to _daemon.exe_.                                                                                         |
| `device_number`               | The device number 0-3 of the drive to mount the image.                                                            |
| `image`                       | The path to the CD-ROM image to mount.                                                                            |
| `wait_seconds_before_program` | How many seconds to wait from the time of mounting the CD-ROM image to the time of running the program.           |
| `program`                     | The path to the program to run, usually the game executable.                                                      |
| `safedisc`                    | Toggle emulation. `On` or `Off` or `Ignore` which leaves the setting as it currently is.                          |
| `securom`                     | Toggle emulation. `On` or `Off` or `Ignore` which leaves the setting as it currently is.                          |
| `laserlock`                   | Toggle emulation. `On` or `Off` or `Ignore` which leaves the setting as it currently is.                          |
| `rmps`                        | Toggle emulation. `On` or `Off` or `Ignore` which leaves the setting as it currently is.                          |
| `unmount`                     | When `true`, unmounts the CD-ROM image after the `program` has finished. Otherwise, `false`.                      |

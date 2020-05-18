# RunAs (Microsoft Windows) - 32bit / 64bit.

![Example](https://i.ibb.co/vdwBk1R/Screenshot-2020-05-18-at-18-11-41.png)

This program is an example about how to easily run any programs as any user.

## Usage

### Mandatory 

* `-u <username>` : Launch program as defined username.
* `-p <password>` : Password associated to username account.
* `-e <program>`  : Executable path (Ex: notepad.exe).

### Optional

* `-d <domain_name>` : Optional domain name (default %USERDOMAIN% environment variable value) for user authentication.
* `-a`<arguments>`   : Arguments to pass to executable.
* `-h`               : Run program without showing window. (Hidden mode)

## Notes

Target program must be accessible by target user otherwise you will receive an Access Denied Error(5).
It is recommended to place target program in a folder accessible by any users such as "C:\ProgramData".




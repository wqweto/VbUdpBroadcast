## VbUdpBroadcast 1.0
UDP Broadcast Forwarder -- a translation to VB6 of the original [WinIPBroadcast](https://github.com/dechamps/WinIPBroadcast) service

### Description

Read original repo [README](https://github.com/dechamps/WinIPBroadcast/blob/master/README.md) for rationale.

### How to use

You can test run `VbUdpBroadcast` as a console application with something like

```
c:> VbUdpBroadcast.exe --console -v
```

... if you don't want to install it as a service at first.

To install `VbUdpBroadcast` service copy executable to a permanent folder and then use

```
c:> VbUdpBroadcast.exe -i
```

### Command-line options

The executable accepts these command-line options:

Option&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | Long&nbsp;Option&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; | Description
------         | ---------         | ------------
`-i`           | `--install`       | Installs `VbUdpBroadcast` NT service.
`-u`           | `--uninstall`     | Stops and removes the `VbUdpBroadcast` NT service.
|              | `--console`       | Starts as a console application (no GUI) with output to console.
`-v`           | `--verbose`       | Verbose output to console.

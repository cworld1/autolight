# `autolight`

Tiny Windows program that turns on/off dark theme at sunrise/sunset.

## Installation

You can download the pre-built binary for the latest release on the
[GitHub releases page](https://github.com/cworld1/autolight/releases). You can put
it anywhere you want on your computer. You can also build from source with
`cargo` if you prefer (see [Building](#building) below).

### Running on startup

You will probably want to run the program at startup. To do so, add a shortcut
to the executable (or the executable itself) in the "Start-up" folder:

1. Copy the executable (`autolight.exe`)
2. `Windows + R` and type `shell:startup`. This will open the "Start-up" folder.
3. In the new explorer window, right-click and select "Paste shortcut". On
   Windows 11, you might have to "Show more options" first.

## Configuration

**The configuration file is mandatory.** It's loaded from
`%USERPROFILE%\.autolight.toml`. This file is constantly watched and changes are
applied automatically.

### Example configuration

#### _%USERPROFILE%/.config/autolight/config.toml_

```toml
# Set this to `true` to terminate the process immediately
disable = false

# Set this to `false` to disable notifications
notifications = true

# Set this to `true` to have dark theme at day and light theme at night
# Why would anyone do this??
invert = false

# How often the program should force a time check (in seconds)
# The program might unsync with the system clock when the computer goes to sleep
# or hibernates. This setting will force a time check to resync it.
refresh_period = 60

# Required for the program to work: replace with YOUR coordinates
[location]
latitude = 48.956775
longitude = 2.463845
```

## Building

`autolight` is written in [Rust](https://www.rust-lang.org/tools/install)!
Install it to build `autolight` from source.

```sh
cargo build
```

You can also build in `release` mode:

```sh
cargo build --release
```

## Usage

Just run the executable. It runs in the background. Everything else happens in
the configuration file.

To stop the process, set `disable = true` in the configuration file.

## License

[MIT](./LICENSE)

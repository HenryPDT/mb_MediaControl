# mb_MediaControl

A MusicBee plugin that adds Windows 10/11 System Media Transport Controls (SMTC) support, so MusicBee shows up in the Windows media overlay, Quick Settings (Win+A) and third-party flyouts such as FluentFlyout and ModernFlyouts.

Based on Steven Mayall's [Windows 10 Media Control Overlay](https://www.getmusicbee.com/addons/plugins/98/windows-10-media-control-overlay/) plugin and the HenryPDT/mb_MediaControl fork.

## Features

- Track metadata: title, artist, album, album artist, track number and artwork
- Windows media controls: play/pause, stop, previous and next track
- Shuffle, repeat mode and timeline (position and seeking) for apps that support them

## What's fixed in 1.1

- **Main cause of failure:** the original obtained SMTC through `BackgroundMediaPlayer`. That API has been deprecated since Windows 10 version 1703 and does not work in regular desktop apps on current Windows builds. Initialisation failed silently, leaving fields null, and disabling the plugin then crashed with a `NullReferenceException` in `SetArtworkThumbnail`.
  SMTC is now obtained with `ISystemMediaTransportControlsInterop.GetForWindow()` for the MusicBee main window, so the media session belongs to the `MusicBee.exe` process itself.
- Every entry point is guarded. A failure no longer leaves null fields behind, and an exception on an SMTC callback thread can no longer take MusicBee down.
- Removed the global keyboard and mouse hook (MouseKeyHook). It only fed a debounce that was disabled (0 ms), so it did nothing.
- One reusable position timer instead of a new one on every play state change (the original leaked timers).
- Artwork is written before `Update()`, so it appears together with the title.
- Errors are written to `mb_MediaControl.log` in MusicBee's data folder (usually `%AppData%\MusicBee`).

## Installation

1. Close MusicBee.
2. Copy `mb_MediaControl.dll` into the `Plugins` folder of your MusicBee installation, replacing any older version.
3. Right-click the DLL → Properties → tick **Unblock** (if the option is shown).
4. Start MusicBee and play something. The session should appear in Quick Settings (Win+A).

## Building

Install the .NET SDK and run in the project folder:

```
dotnet build -c Release
```

Output: `bin\Release\net48\mb_MediaControl.dll`.

## Acknowledgments

- [Windows 10 Media Control Overlay](https://www.getmusicbee.com/addons/plugins/98/windows-10-media-control-overlay/) by Steven Mayall
- [CharlieJiang version](https://getmusicbee.com/forum/index.php?topic=21240.45)
- [HenryPDT/mb_MediaControl](https://github.com/HenryPDT/mb_MediaControl)
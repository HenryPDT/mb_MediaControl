# mb_MediaControl (album art fix)

This is a fork of [HenryPDT/mb_MediaControl](https://github.com/HenryPDT/mb_MediaControl) that fixes
album art sometimes being missing in the Windows media overlay and in apps that read it, such as
[Music Presence](https://musicpresence.app) for Discord.

**[Download the latest release](https://github.com/YOUR-USERNAME/mb_MediaControl/releases/latest)**

## What's fixed
- The cover is now set before the track is published (it was a race condition)
- Large embedded covers are downscaled to 600 px, so they load instantly
- Covers that MusicBee finds late (downloaded or slow to load) are picked up
- Covers are re-sent shortly after a track change, in case Windows drops one
- A leaking position timer and a crash on close were fixed

## Install
1. Close MusicBee.
2. Put `mb_MediaControl.dll` in your MusicBee `Plugins` folder, replacing the old one if you have it.
3. Start MusicBee. **Edit → Preferences → Plugins** should show *Media Control 1.0.5*.

If the plugin doesn't load, right-click the DLL, choose **Properties**, tick **Unblock**, and restart MusicBee.

## Known limitation
Skipping through tracks very quickly can leave Discord a few seconds behind. This also happens with the
original plugin and comes from Music Presence's update limits, not from this plugin.

## Building
Open `MediaControl.sln` in Visual Studio (with the *.NET desktop development* workload), set the
configuration to **Release**, and build. The DLL is in `bin\Release`. Debug builds also write a
diagnostic log to `%TEMP%\mb_MediaControl.log`; Release builds don't.

## Credits
Original plugin by its author and contributors (see the upstream repository). This fork only changes
the album art handling.

---

(original README below)
# mb_MediaControl
This plugin adds SMTC (Media Control Overlay) support feature to MusicBee for Windows 10/11

## Features
- Media metadata display including title and artist information
- Support Windows media controls
  - Play/Pause
  - Previous Track
  - Next Track
- Support [ModernFlyout](https://github.com/ModernFlyouts-Community/ModernFlyouts) media controls
  - Shuffle
  - Repeat Mode
  - Timeline Info and Control

## Usage
Put `mb_MediaControl.dll` into `MusicBee/Plugins` folder

## Acknowledgments
* [Windows 10 Media Control Overlay](https://www.getmusicbee.com/addons/plugins/98/windows-10-media-control-overlay/)
* [CharlieJiang version](https://getmusicbee.com/forum/index.php?topic=21240.45)

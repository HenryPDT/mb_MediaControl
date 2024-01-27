using System;
using System.Threading;
using System.Runtime.InteropServices;
using System.Collections.Generic;
using Windows.Media;
using Windows.Media.Core;
using Windows.Media.Playback;
using Windows.Storage;
using Windows.Storage.Streams;
using static System.IO.WindowsRuntimeStreamExtensions;
using System.IO;
using System.Runtime.InteropServices.WindowsRuntime;
using Gma.System.MouseKeyHook;
using System.Windows.Forms;
using Windows.Media.Control;

namespace MusicBeePlugin
{
    public partial class Plugin
    {
        private MusicBeeApiInterface mbApiInterface;
        private readonly PluginInfo about = new PluginInfo();
        private SystemMediaTransportControls systemMediaControls;
        private SystemMediaTransportControlsDisplayUpdater displayUpdater;
        private MusicDisplayProperties musicProperties;
        private InMemoryRandomAccessStream artworkStream;

        private IKeyboardMouseEvents globalHook;
        // Disable this...
        private int mediaKeysInvalidateBeforeMs = 0; // media control buttons won't trigger for 2000 ms after media button is pressed
        private DateTime lastPlayPauseKeyPress;
        private DateTime lastStopKeyPress;
        private DateTime lastPreviousTrackKeyPress;
        private DateTime lastNextTrackKeyPress;

        private bool trackChangeListenerDisabled = false;
        private System.Threading.Timer timer;

        public PluginInfo Initialise(IntPtr apiInterfacePtr)
        {
            SubscribeGlobalHooks();
            mbApiInterface = new MusicBeeApiInterface();
            mbApiInterface.Initialise(apiInterfacePtr);
            about.PluginInfoVersion = PluginInfoVersion;
            about.Name = "Media Control";
            about.Description = "Enables MusicBee to interact with the Windows 10/11 Media Control overlay.";
            about.Author = "Steven Mayall";
            about.TargetApplication = "";   //  the name of a Plugin Storage device or panel header for a dockable panel
            about.Type = PluginType.General;
            about.VersionMajor = 1;  // your plugin version
            about.VersionMinor = 0;
            about.Revision = 4;
            about.MinInterfaceVersion = MinInterfaceVersion;
            about.MinApiRevision = MinApiRevision;
            about.ReceiveNotifications = (ReceiveNotificationFlags.PlayerEvents | ReceiveNotificationFlags.TagEvents);
            about.ConfigurationPanelHeight = 0;   // height in pixels that musicbee should reserve in a panel for config settings. When set, a handle to an empty panel will be passed to the Configure function
            return about;
        }

        public bool Configure(IntPtr panelHandle)
        {
            return false;
        }

        // called by MusicBee when the user clicks Apply or Save in the MusicBee Preferences screen.
        // its up to you to figure out whether anything has changed and needs updating
        public void SaveSettings()
        {
        }

        // MusicBee is closing the plugin (plugin is being disabled by user or MusicBee is shutting down)
        public void Close(PluginCloseReason reason)
        {
            UnsubscribeGlobalHooks();
            SetArtworkThumbnail(null);
            timer.Dispose();
        }

        // uninstall this plugin - clean up any persisted files
        public void Uninstall()
        {
        }

        private void MediaControl_PlayPauseButtonPress(bool pause)
        {
            trackChangeListenerDisabled = true;
            try
            {
                if (DateTime.Now.Subtract(lastPlayPauseKeyPress).TotalMilliseconds > mediaKeysInvalidateBeforeMs)
                {
                    var state = mbApiInterface.Player_GetPlayState();
                    switch (state)
                    {
                        case PlayState.Playing:
                            if (pause)
                                mbApiInterface.Player_PlayPause();
                            break;
                        case PlayState.Paused:
                            if (!pause)
                                mbApiInterface.Player_PlayPause();
                            break;
                        case PlayState.Stopped:
                        case PlayState.Loading:
                        case PlayState.Undefined:
                        default:
                            break; // Ignored
                    }
                }
                SetPlayerState();
            }
            finally
            {
                trackChangeListenerDisabled = false; // atomic!
            }
        }

        private void MediaControl_StopButtonPress()
        {
            trackChangeListenerDisabled = true;
            try
            {
                if (DateTime.Now.Subtract(lastStopKeyPress).TotalMilliseconds > mediaKeysInvalidateBeforeMs)
                    mbApiInterface.Player_Stop();
                SetPlayerState();
            }
            finally
            {
                trackChangeListenerDisabled = false; // atomic!
            }
        }

        private void MediaControl_PreviousTrackButtonPress()
        {
            trackChangeListenerDisabled = true;
            try
            {
                if (DateTime.Now.Subtract(lastPreviousTrackKeyPress).TotalMilliseconds > mediaKeysInvalidateBeforeMs)
                    mbApiInterface.Player_PlayPreviousTrack();
                SetDisplayValues();
            }
            finally
            {
                trackChangeListenerDisabled = false; // atomic!
            }
        }

        private void MediaControl_NextTrackButtonPress()
        {
            trackChangeListenerDisabled = true;
            try
            {
                if (DateTime.Now.Subtract(lastNextTrackKeyPress).TotalMilliseconds > mediaKeysInvalidateBeforeMs)
                    mbApiInterface.Player_PlayNextTrack();
                SetDisplayValues();
            }
            finally
            {
                trackChangeListenerDisabled = false; // atomic!
            }
        }

        // receive event notifications from MusicBee
        // you need to set about.ReceiveNotificationFlags = PlayerEvents to receive all notifications, and not just the startup event
        public void ReceiveNotification(string sourceFileUrl, NotificationType type)
        {
            switch (type)
            {
                case NotificationType.PluginStartup:
                    systemMediaControls = BackgroundMediaPlayer.Current.SystemMediaTransportControls;
                    systemMediaControls.PlaybackStatus = MediaPlaybackStatus.Closed;
                    systemMediaControls.IsEnabled = true;
                    systemMediaControls.IsPlayEnabled = true;
                    systemMediaControls.IsPauseEnabled = true;
                    systemMediaControls.IsStopEnabled = true;
                    systemMediaControls.IsPreviousEnabled = true;
                    systemMediaControls.IsNextEnabled = true;
                    systemMediaControls.IsRewindEnabled = false;
                    systemMediaControls.IsFastForwardEnabled = false;
                    systemMediaControls.ButtonPressed += SystemMediaControls_ButtonPressed;
                    systemMediaControls.PlaybackPositionChangeRequested += SystemMediaControls_PlaybackPositionChangeRequested;
                    systemMediaControls.PlaybackRateChangeRequested += SystemMediaControls_PlaybackRateChangeRequested;
                    systemMediaControls.ShuffleEnabledChangeRequested += SystemMediaControls_ShuffleEnabledChangeRequested;
                    systemMediaControls.AutoRepeatModeChangeRequested += SystemMediaControls_AutoRepeatModeChangeRequested;
                    displayUpdater = systemMediaControls.DisplayUpdater;
                    displayUpdater.Type = MediaPlaybackType.Music;
                    musicProperties = displayUpdater.MusicProperties;
                    SetDisplayValues();
                    SetShuffleState();
                    SetRepeatState();
                    break;
                case NotificationType.PlayStateChanged:
                    if (!trackChangeListenerDisabled)
                        SetPlayerState();
                    break;
                case NotificationType.TrackChanged:
                    if (!trackChangeListenerDisabled)
                        SetDisplayValues();
                    break;
                case NotificationType.PlayerShuffleChanged:
                    if (!trackChangeListenerDisabled)
                        SetShuffleState();
                    break;
                case NotificationType.PlayerRepeatChanged:
                    if (!trackChangeListenerDisabled)
                        SetRepeatState();
                    break;
            }


        }

        private void SystemMediaControls_ButtonPressed(SystemMediaTransportControls smtc, SystemMediaTransportControlsButtonPressedEventArgs args)
        {
            switch (args.Button)
            {
                case SystemMediaTransportControlsButton.Stop:
                    MediaControl_StopButtonPress();
                    break;
                case SystemMediaTransportControlsButton.Play:
                    MediaControl_PlayPauseButtonPress(false);
                    break;
                case SystemMediaTransportControlsButton.Pause:
                    MediaControl_PlayPauseButtonPress(true);
                    break;
                case SystemMediaTransportControlsButton.Next:
                    MediaControl_NextTrackButtonPress();
                    break;
                case SystemMediaTransportControlsButton.Previous:
                    MediaControl_PreviousTrackButtonPress();
                    break;
                case SystemMediaTransportControlsButton.Rewind:
                    break;
                case SystemMediaTransportControlsButton.FastForward:
                    break;
                case SystemMediaTransportControlsButton.ChannelUp:
                    mbApiInterface.Player_SetVolume(mbApiInterface.Player_GetVolume() + 0.05F);
                    break;
                case SystemMediaTransportControlsButton.ChannelDown:
                    mbApiInterface.Player_SetVolume(mbApiInterface.Player_GetVolume() - 0.05F);
                    break;
            }
        }

        private void SystemMediaControls_PlaybackPositionChangeRequested(SystemMediaTransportControls smtc, PlaybackPositionChangeRequestedEventArgs args)
        {
            mbApiInterface.Player_SetPosition((int)args.RequestedPlaybackPosition.TotalMilliseconds);

        }

        private void SystemMediaControls_PlaybackRateChangeRequested(SystemMediaTransportControls smtc, PlaybackRateChangeRequestedEventArgs args)
        {
        }

        private void SystemMediaControls_AutoRepeatModeChangeRequested(SystemMediaTransportControls smtc, AutoRepeatModeChangeRequestedEventArgs args)
        {
            switch (args.RequestedAutoRepeatMode)
            {
                case MediaPlaybackAutoRepeatMode.Track:
                    mbApiInterface.Player_SetRepeat(RepeatMode.One);
                    break;
                case MediaPlaybackAutoRepeatMode.List:
                    mbApiInterface.Player_SetRepeat(RepeatMode.All);
                    break;
                case MediaPlaybackAutoRepeatMode.None:
                    mbApiInterface.Player_SetRepeat(RepeatMode.None);
                    break;
            }
        }

        private void SystemMediaControls_ShuffleEnabledChangeRequested(SystemMediaTransportControls smtc, ShuffleEnabledChangeRequestedEventArgs args)
        {
            mbApiInterface.Player_SetShuffle(args.RequestedShuffleEnabled);

        }

        private void SetDisplayValues()
        {
            displayUpdater.ClearAll();
            displayUpdater.Type = MediaPlaybackType.Music;
            SetArtworkThumbnail(null);
            var url = mbApiInterface.NowPlaying_GetFileUrl();
            if (url != null)
            {
                musicProperties.AlbumArtist = mbApiInterface.NowPlaying_GetFileTag(MetaDataType.AlbumArtist);
                musicProperties.AlbumTitle = mbApiInterface.NowPlaying_GetFileTag(MetaDataType.Album);
                if (uint.TryParse(mbApiInterface.NowPlaying_GetFileTag(MetaDataType.TrackCount), out var value))
                    musicProperties.AlbumTrackCount = value;
                musicProperties.Artist = mbApiInterface.NowPlaying_GetFileTag(MetaDataType.Artist);
                musicProperties.Title = mbApiInterface.NowPlaying_GetFileTag(MetaDataType.TrackTitle);
                if (string.IsNullOrEmpty(musicProperties.Title))
                    musicProperties.Title = url.Substring(url.LastIndexOfAny(new char[] { '/', '\\' }) + 1);
                if (uint.TryParse(mbApiInterface.NowPlaying_GetFileTag(MetaDataType.TrackNo), out value))
                    musicProperties.TrackNumber = value;
                //musicProperties.Genres = mbApiInterface.NowPlaying_GetFileTag(MetaDataType.Genres).Split(new string[] {"; "}, StringSplitOptions.RemoveEmptyEntries);
                mbApiInterface.Library_GetArtworkEx(url, 0, true, out _, out _, out var imageData);
                SetArtworkThumbnail(imageData);
            }
            displayUpdater.Update();
        }

        private void SetPlayerState()
        {
            switch (mbApiInterface.Player_GetPlayState())
            {
                case PlayState.Playing:
                    systemMediaControls.PlaybackStatus = MediaPlaybackStatus.Playing;
                    timer = new System.Threading.Timer(SetPositionState, null, 0, 1000);
                    break;
                case PlayState.Paused:
                    systemMediaControls.PlaybackStatus = MediaPlaybackStatus.Paused;
                    timer.Dispose();
                    break;
                case PlayState.Stopped:
                    systemMediaControls.PlaybackStatus = MediaPlaybackStatus.Stopped;
                    timer.Dispose();
                    break;
            }
        }

        private void SetShuffleState()
        {
            systemMediaControls.ShuffleEnabled = mbApiInterface.Player_GetShuffle();
        }

        private void SetRepeatState()
        {
            switch (mbApiInterface.Player_GetRepeat())
            {
                case RepeatMode.One:
                    systemMediaControls.AutoRepeatMode = MediaPlaybackAutoRepeatMode.Track;
                    break;
                case RepeatMode.All:
                    systemMediaControls.AutoRepeatMode = MediaPlaybackAutoRepeatMode.List;
                    break;
                case RepeatMode.None:
                    systemMediaControls.AutoRepeatMode = MediaPlaybackAutoRepeatMode.None;
                    break;
            }
        }

        private void SetPositionState(object state)
        {
            var timelineProperties = new SystemMediaTransportControlsTimelineProperties();

            timelineProperties.StartTime = TimeSpan.FromSeconds(0);
            timelineProperties.MinSeekTime = TimeSpan.FromSeconds(0);
            timelineProperties.Position = TimeSpan.FromMilliseconds(mbApiInterface.Player_GetPosition());
            timelineProperties.MaxSeekTime = TimeSpan.FromMilliseconds(mbApiInterface.NowPlaying_GetDuration());
            timelineProperties.EndTime = TimeSpan.FromMilliseconds(mbApiInterface.NowPlaying_GetDuration());

            systemMediaControls.UpdateTimelineProperties(timelineProperties);
        }

        private async void SetArtworkThumbnail(byte[] data)
        {
            if (artworkStream != null)
                artworkStream.Dispose();
            if (data == null)
            {
                artworkStream = null;
                displayUpdater.Thumbnail = null;
            }
            else
            {
                new MemoryStream(data).AsInputStream();

                artworkStream = new InMemoryRandomAccessStream();
                await artworkStream.WriteAsync(data.AsBuffer());
                displayUpdater.Thumbnail = RandomAccessStreamReference.CreateFromStream(artworkStream);
            }
        }

        private void SubscribeGlobalHooks()
        {
            globalHook = Hook.GlobalEvents();
            globalHook.KeyPress += GlobalHook_KeyPress;
        }

        private void UnsubscribeGlobalHooks()
        {
            globalHook.KeyPress -= GlobalHook_KeyPress;
            globalHook.Dispose();
        }

        private void GlobalHook_KeyPress(object sender, KeyPressEventArgs e)
        {
            switch ((Keys)e.KeyChar)
            {
                case Keys.MediaPlayPause:
                    lastPlayPauseKeyPress = DateTime.Now;
                    break;
                case Keys.MediaStop:
                    lastStopKeyPress = DateTime.Now;
                    break;
                case Keys.MediaPreviousTrack:
                    lastPreviousTrackKeyPress = DateTime.Now;
                    break;
                case Keys.MediaNextTrack:
                    lastNextTrackKeyPress = DateTime.Now;
                    break;
                default:
                    break;
            }
        }

    }

}
// 'ex' is only read by Log(), which Release builds compile away
#pragma warning disable CS0168
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
        private IRandomAccessStream artworkStream;

        private IKeyboardMouseEvents globalHook;
        // Disable this...
        private int mediaKeysInvalidateBeforeMs = 0; // media control buttons won't trigger for 2000 ms after media button is pressed
        private DateTime lastPlayPauseKeyPress;
        private DateTime lastStopKeyPress;
        private DateTime lastPreviousTrackKeyPress;
        private DateTime lastNextTrackKeyPress;

        private bool trackChangeListenerDisabled = false;
        private System.Threading.Timer timer;
        // NOTE: MusicBee calls the plugin on a thread without a message loop, so WinForms timers would never
        // fire. Plain threading timers are used instead.
        private System.Threading.Timer artworkRetryTimer; // re-checks artwork that wasn't ready yet
        private int artworkRetryCount;
        private System.Threading.Timer reassertTimer; // re-sends the cover shortly after a track change
        private StorageFile currentArtworkFile;
        private int reassertGeneration;
        private string lastPublishedUrl; // file the Windows media controls are currently showing
        private readonly object smtcLock = new object();
        private int artworkFileCounter;

        // Diagnostic log: compiled into Debug builds only. Release builds contain no logging code at all
        // (calls to a [Conditional] method are removed by the compiler). File: %TEMP%\mb_MediaControl.log
        private static readonly string logPath = Path.Combine(Path.GetTempPath(), "mb_MediaControl.log");

        [System.Diagnostics.Conditional("DEBUG")]
        private static void Log(string msg)
        {
            try
            {
                var fi = new FileInfo(logPath);
                if (fi.Exists && fi.Length > 512 * 1024) fi.Delete(); // keep it small
                File.AppendAllText(logPath, DateTime.Now.ToString("HH:mm:ss.fff") + " [t" + Thread.CurrentThread.ManagedThreadId + "] " + msg + Environment.NewLine);
            }
            catch { }
        }

        // cover images handed to Windows are stored here (small files, cleaned up automatically)
        private static string ArtworkDir { get { return Path.Combine(Path.GetTempPath(), "mb_MediaControl_art"); } }

        public PluginInfo Initialise(IntPtr apiInterfacePtr)
        {
            Log("Plugin loaded, log path: " + logPath);
            SubscribeGlobalHooks();
            mbApiInterface = new MusicBeeApiInterface();
            mbApiInterface.Initialise(apiInterfacePtr);
            LoadSettings();
            about.PluginInfoVersion = PluginInfoVersion;
            about.Name = "Media Control";
            about.Description = "Enables MusicBee to interact with the Windows 10/11 Media Control overlay.";
            about.Author = "Steven Mayall";
            about.TargetApplication = "";   //  the name of a Plugin Storage device or panel header for a dockable panel
            about.Type = PluginType.General;
            about.VersionMajor = 1;  // your plugin version
            about.VersionMinor = 0;
            about.Revision = 6;
            about.MinInterfaceVersion = MinInterfaceVersion;
            about.MinApiRevision = MinApiRevision;
            about.ReceiveNotifications = (ReceiveNotificationFlags.PlayerEvents | ReceiveNotificationFlags.TagEvents | ReceiveNotificationFlags.DownloadEvents);
            about.ConfigurationPanelHeight = 30;   // height in pixels that musicbee should reserve in a panel for config settings. When set, a handle to an empty panel will be passed to the Configure function
            return about;
        }

        // ---- settings -------------------------------------------------------------------------------
        // Largest side (in pixels) a cover may have before it is shrunk for the Windows media controls.
        // 0 = never shrink. Larger values give sharper covers (e.g. for a self-hosted Music Presence
        // cover proxy) but are slower for Windows to load and for presence tools to upload.
        private const int DefaultMaxArtworkSide = 600;
        private const int MaxAllowedArtworkSide = 4096;
        private volatile int maxArtworkSide = DefaultMaxArtworkSide;
        private NumericUpDown sizeBox;

        private string SettingsPath
        {
            get { return Path.Combine(mbApiInterface.Setting_GetPersistentStoragePath(), "mb_MediaControl.ini"); }
        }

        private void LoadSettings()
        {
            try
            {
                if (!File.Exists(SettingsPath)) return;
                foreach (var line in File.ReadAllLines(SettingsPath))
                {
                    var parts = line.Split(new[] { '=' }, 2);
                    if (parts.Length == 2 && parts[0].Trim() == "MaxArtworkSide" && int.TryParse(parts[1].Trim(), out var v))
                        maxArtworkSide = Math.Max(0, Math.Min(MaxAllowedArtworkSide, v));
                }
            }
            catch { } // unreadable settings: keep the default
            Log("Max cover size: " + maxArtworkSide);
        }

        public bool Configure(IntPtr panelHandle)
        {
            if (panelHandle != IntPtr.Zero)
            {
                var panel = (Panel)Panel.FromHandle(panelHandle);
                var label = new Label
                {
                    AutoSize = true,
                    Location = new System.Drawing.Point(0, 6),
                    Text = "Max cover size in pixels (0 = don't shrink; default " + DefaultMaxArtworkSide + "):"
                };
                sizeBox = new NumericUpDown
                {
                    Minimum = 0,
                    Maximum = MaxAllowedArtworkSide,
                    Increment = 100,
                    Value = maxArtworkSide,
                    Width = 70
                };
                panel.Controls.Add(label);
                sizeBox.Location = new System.Drawing.Point(label.PreferredWidth + 8, 3);
                panel.Controls.Add(sizeBox);
            }
            return false;
        }

        // called by MusicBee when the user clicks Apply or Save in the MusicBee Preferences screen.
        public void SaveSettings()
        {
            if (sizeBox == null || sizeBox.IsDisposed) return;
            int value = (int)sizeBox.Value;
            if (value == maxArtworkSide) return;
            maxArtworkSide = value;
            try { File.WriteAllText(SettingsPath, "MaxArtworkSide=" + value + Environment.NewLine); }
            catch { }
            lock (shrinkCache) { shrinkCache.Clear(); }
            // show the playing track's cover at the new size right away
            if (displayUpdater != null)
            {
                try { RefreshArtworkOnly(); } catch { }
            }
        }

        // MusicBee is closing the plugin (plugin is being disabled by user or MusicBee is shutting down)
        public void Close(PluginCloseReason reason)
        {
            UnsubscribeGlobalHooks();
            SetArtworkThumbnail(null);
            timer?.Dispose();
            artworkRetryTimer?.Dispose();
            reassertTimer?.Dispose();
            CleanOldArtworkFiles(ArtworkDir, 0);
        }

        // uninstall this plugin - clean up any persisted files
        public void Uninstall()
        {
            try { File.Delete(SettingsPath); } catch { }
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
            Log("Previous button pressed (media key / Windows overlay)");
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
            Log("Next button pressed (media key / Windows overlay)");
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
                    CleanOldArtworkFiles(ArtworkDir, 0); // leftovers from a previous session
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
                    Log("TrackChanged notification (ignored=" + trackChangeListenerDisabled + ")");
                    if (!trackChangeListenerDisabled)
                        SetDisplayValues();
                    break;
                case NotificationType.NowPlayingArtworkReady:
                    Log("NowPlayingArtworkReady notification");
                    // MusicBee finished locating/downloading artwork for the playing track
                    RefreshArtworkOnly();
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
            lock (smtcLock) { SetDisplayValuesCore(); }
        }

        private void SetDisplayValuesCore()
        {
            Interlocked.Increment(ref reassertGeneration); // cancels re-sends belonging to the previous track
            currentArtworkFile = null;
            displayUpdater.ClearAll();
            displayUpdater.Type = MediaPlaybackType.Music;
            var url = mbApiInterface.NowPlaying_GetFileUrl();
            lastPublishedUrl = url;
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
                // The cover must be set synchronously and BEFORE displayUpdater.Update() (see SetArtworkThumbnail)
                Log("SetDisplayValues: " + musicProperties.Artist + " - " + musicProperties.Title);
                var imageData = GetCurrentArtwork(url);
                Log("  artwork result: " + (imageData == null ? "NONE" : imageData.Length + " bytes, header " + BitConverter.ToString(imageData, 0, Math.Min(4, imageData.Length))));
                SetArtworkThumbnail(imageData);
                if (imageData == null)
                    StartArtworkRetry(); // artwork may simply not be ready yet
                else
                    StopArtworkRetry();
                StartReassert(); // verifies the display and re-sends the cover shortly after (see ScheduleReassert)
            }
            displayUpdater.Update();
        }

        // Tries every source MusicBee offers: embedded/folder art from the library first,
        // then whatever MusicBee is currently displaying (including downloaded artwork).
        private byte[] GetCurrentArtwork(string url)
        {
            try
            {
                var ok = mbApiInterface.Library_GetArtworkEx(url, 0, true, out var loc, out var picUrl, out var data);
                Log("Library_GetArtworkEx ok=" + ok + " loc=" + loc + " url=" + picUrl + " bytes=" + (data == null ? "null" : data.Length.ToString()));
                if (data != null && data.Length > 0)
                    return data;
            }
            catch (Exception ex) { Log("Library_GetArtworkEx threw " + ex.Message); }
            try
            {
                var b64 = mbApiInterface.NowPlaying_GetArtwork();
                Log("NowPlaying_GetArtwork len=" + (b64 == null ? "null" : b64.Length.ToString()));
                if (!string.IsNullOrEmpty(b64))
                    return Convert.FromBase64String(b64);
            }
            catch (Exception ex) { Log("NowPlaying_GetArtwork threw " + ex.Message); }
            try
            {
                var b64 = mbApiInterface.NowPlaying_GetDownloadedArtwork();
                Log("NowPlaying_GetDownloadedArtwork len=" + (b64 == null ? "null" : b64.Length.ToString()));
                if (!string.IsNullOrEmpty(b64))
                    return Convert.FromBase64String(b64);
            }
            catch (Exception ex) { Log("NowPlaying_GetDownloadedArtwork threw " + ex.Message); }
            return null;
        }

        // Updates only the cover, without clearing the rest of the metadata
        private void RefreshArtworkOnly()
        {
            if (displayUpdater == null) return;
            lock (smtcLock)
            {
                var url = mbApiInterface.NowPlaying_GetFileUrl();
                if (url == null) return;
                var imageData = GetCurrentArtwork(url);
                if (imageData == null) return;
                StopArtworkRetry();
                Log("RefreshArtworkOnly: got " + imageData.Length + " bytes");
                SetArtworkThumbnail(imageData);
                displayUpdater.Update();
                StartReassert();
            }
        }

        private void StopArtworkRetry()
        {
            artworkRetryTimer?.Dispose();
            artworkRetryTimer = null;
        }

        private void StartArtworkRetry()
        {
            StopArtworkRetry();
            artworkRetryCount = 0;
            artworkRetryTimer = new System.Threading.Timer(_ =>
            {
                artworkRetryCount++;
                try { RefreshArtworkOnly(); } catch (Exception ex) { Log("retry failed: " + ex.Message); }
                if (artworkRetryCount >= 5)
                    StopArtworkRetry(); // give up after ~5 seconds; no artwork exists
            }, null, 1000, 1000);
        }

        // Shortly after each track change: check that Windows shows the track MusicBee is really playing,
        // and re-send the cover (see ScheduleReassert).
        private void StartReassert()
        {
            ScheduleReassert(1000, 1, Interlocked.Increment(ref reassertGeneration));
        }

        private void ScheduleReassert(int delayMs, int attempt, int gen)
        {
            if (gen != reassertGeneration) return; // a newer track has started
            reassertTimer?.Dispose();
            reassertTimer = new System.Threading.Timer(_ =>
            {
                try
                {
                    if (gen != reassertGeneration) return;
                    if (VerifyDisplayedTrack()) return; // display was stale; it was republished (new cycle started)
                    var resent = false;
                    lock (smtcLock)
                    {
                        if (gen != reassertGeneration) return;
                        if (currentArtworkFile != null)
                        {
                            // fresh reference to the same cover file, then send the display state again
                            displayUpdater.Thumbnail = RandomAccessStreamReference.CreateFromFile(currentArtworkFile);
                            displayUpdater.Update();
                            resent = true;
                            Log("  re-sent cover (attempt " + attempt + ")");
                        }
                    }
                    // Two re-sends per track (~1 s and ~3 s after the change). Windows occasionally drops a cover
                    // that was set right at the track change, even though it was handed over correctly.
                    if (resent && attempt < 2)
                        ScheduleReassert(2000, attempt + 1, gen);
                }
                catch (Exception ex) { Log("  re-send failed: " + ex.Message); }
            }, null, delayMs, Timeout.Infinite);
        }

        // If MusicBee is playing a different file than the one we published (e.g. a skip arrived while we were
        // reading the "now playing" info), publish again. Returns true if it had to.
        private bool VerifyDisplayedTrack()
        {
            string url;
            try { url = mbApiInterface.NowPlaying_GetFileUrl(); }
            catch { return false; }
            if (url == null || url == lastPublishedUrl) return false;
            Log("  stale display (showing '" + lastPublishedUrl + "', playing '" + url + "'), republishing");
            SetDisplayValues();
            return true;
        }

        private void SetPlayerState()
        {
            switch (mbApiInterface.Player_GetPlayState())
            {
                case PlayState.Playing:
                    systemMediaControls.PlaybackStatus = MediaPlaybackStatus.Playing;
                    timer?.Dispose(); // don't leak a timer on every play/resume
                    timer = new System.Threading.Timer(SetPositionState, null, 0, 1000);
                    break;
                case PlayState.Paused:
                    systemMediaControls.PlaybackStatus = MediaPlaybackStatus.Paused;
                    timer?.Dispose();
                    break;
                case PlayState.Stopped:
                    systemMediaControls.PlaybackStatus = MediaPlaybackStatus.Stopped;
                    timer?.Dispose();
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

        // Hands the cover to Windows as a small file:
        //  - Windows can re-read a file whenever an app asks (the overlay, Discord presence tools, ...),
        //    which a one-shot in-memory stream does not reliably allow;
        //  - it is done synchronously, so the thumbnail is in place before displayUpdater.Update() runs
        //    (the previous async version could publish the track before the cover was assigned).
        // An in-memory stream is only used as a fallback if writing the file fails.
        private void SetArtworkThumbnail(byte[] data)
        {
            if (data == null || data.Length == 0)
            {
                artworkStream = null;
                currentArtworkFile = null;
                displayUpdater.Thumbnail = null;
                return;
            }

            data = ShrinkArtwork(data, maxArtworkSide);

            try
            {
                var dir = ArtworkDir;
                Directory.CreateDirectory(dir);
                // a fresh file name every time, so a file Windows still has open is never overwritten
                var path = Path.Combine(dir, "cover_" + (++artworkFileCounter) + "_" + DateTime.Now.Ticks + ArtworkExtension(data));
                File.WriteAllBytes(path, data);
                var file = StorageFile.GetFileFromPathAsync(path).AsTask().GetAwaiter().GetResult();
                currentArtworkFile = file;
                displayUpdater.Thumbnail = RandomAccessStreamReference.CreateFromFile(file);
                Log("  thumbnail set from file " + path);
                CleanOldArtworkFiles(dir, 8);
                return;
            }
            catch (Exception ex)
            {
                Log("  file thumbnail failed (" + ex.Message + "), falling back to memory stream");
            }

            currentArtworkFile = null;
            var stream = new InMemoryRandomAccessStream();
            stream.WriteAsync(data.AsBuffer()).AsTask().Wait();
            stream.FlushAsync().AsTask().Wait();
            stream.Seek(0);
            artworkStream = stream; // keep a reference so it stays alive
            displayUpdater.Thumbnail = RandomAccessStreamReference.CreateFromStream(stream);
            Log("  thumbnail set from memory stream");
        }

        // Embedded covers can be several MB. Windows has to load them before any app can read them and
        // presence tools have to upload them, so big files make cover loading slow and unreliable. A 600 px JPEG
        // is plenty for the Windows media overlay and for Discord. The limit is configurable (see settings above).
        private static readonly Dictionary<string, byte[]> shrinkCache = new Dictionary<string, byte[]>();

        private static byte[] ShrinkArtwork(byte[] data, int maxSide)
        {
            if (maxSide <= 0) return data; // shrinking disabled
            // tracks of one album share a cover: remember recent results instead of decoding the image again
            string key = null;
            try
            {
                using (var md5 = System.Security.Cryptography.MD5.Create()) // just a cache key, not security
                    key = maxSide + ":" + BitConverter.ToString(md5.ComputeHash(data));
            }
            catch { } // e.g. MD5 blocked by FIPS policy: simply no caching

            if (key != null)
            {
                lock (shrinkCache)
                {
                    if (shrinkCache.TryGetValue(key, out var cached))
                    {
                        Log("  cover from cache (" + cached.Length + " bytes)");
                        return cached;
                    }
                }
            }
            var result = ShrinkArtworkCore(data, maxSide);
            if (key != null)
            {
                lock (shrinkCache)
                {
                    if (shrinkCache.Count >= 24) shrinkCache.Clear();
                    shrinkCache[key] = result;
                }
            }
            return result;
        }

        private static byte[] ShrinkArtworkCore(byte[] data, int maxSide)
        {
            try
            {
                using (var input = new MemoryStream(data))
                using (var img = System.Drawing.Image.FromStream(input))
                {
                    int w = img.Width, h = img.Height;
                    if (w <= maxSide && h <= maxSide && data.Length < 300 * 1024)
                        return data; // already small
                    double scale = Math.Min(1.0, (double)maxSide / Math.Max(w, h));
                    int nw = Math.Max(1, (int)Math.Round(w * scale));
                    int nh = Math.Max(1, (int)Math.Round(h * scale));
                    using (var bmp = new System.Drawing.Bitmap(nw, nh, System.Drawing.Imaging.PixelFormat.Format24bppRgb))
                    {
                        using (var g = System.Drawing.Graphics.FromImage(bmp))
                        {
                            g.Clear(System.Drawing.Color.White);
                            g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                            g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                            g.DrawImage(img, 0, 0, nw, nh);
                        }
                        System.Drawing.Imaging.ImageCodecInfo jpeg = null;
                        foreach (var c in System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders())
                            if (c.MimeType == "image/jpeg") { jpeg = c; break; }
                        if (jpeg == null) return data;
                        using (var ms = new MemoryStream())
                        using (var ep = new System.Drawing.Imaging.EncoderParameters(1))
                        {
                            ep.Param[0] = new System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 85L);
                            bmp.Save(ms, jpeg, ep);
                            var result = ms.ToArray();
                            Log("  shrunk cover " + w + "x" + h + " (" + data.Length + " bytes) -> " + nw + "x" + nh + " (" + result.Length + " bytes)");
                            return result.Length < data.Length ? result : data;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("  shrink failed (" + ex.Message + "), using original");
                return data;
            }
        }

        private static string ArtworkExtension(byte[] d)
        {
            if (d.Length > 3 && d[0] == 0x89 && d[1] == 0x50) return ".png";
            if (d.Length > 3 && d[0] == 0x47 && d[1] == 0x49) return ".gif";
            if (d.Length > 2 && d[0] == 0x42 && d[1] == 0x4D) return ".bmp";
            if (d.Length > 11 && d[0] == 0x52 && d[1] == 0x49 && d[8] == 0x57) return ".webp";
            return ".jpg";
        }

        // keeps only the newest `keep` cover files (0 = delete them all); files still in use are skipped
        private static void CleanOldArtworkFiles(string dir, int keep)
        {
            try
            {
                if (!Directory.Exists(dir)) return;
                var files = new DirectoryInfo(dir).GetFiles();
                Array.Sort(files, (x, y) => y.CreationTimeUtc.CompareTo(x.CreationTimeUtc));
                for (int i = keep; i < files.Length; i++)
                {
                    try { files[i].Delete(); } catch { }
                }
            }
            catch { }
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

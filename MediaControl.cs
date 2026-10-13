// mb_MediaControl - Windows 10/11 media controls (SMTC) for MusicBee.
//
// Based on the "Windows 10 Media Control Overlay" plugin by Steven Mayall and the
// HenryPDT/mb_MediaControl fork. Patched version 1.1:
//   * SMTC is obtained with ISystemMediaTransportControlsInterop.GetForWindow() on the
//     MusicBee main window instead of the deprecated BackgroundMediaPlayer, which does
//     not work for desktop apps on current Windows builds. The media session therefore
//     belongs to MusicBee.exe itself.
//   * Every entry point is guarded: a failed initialisation no longer leaves null fields
//     behind that crash later (e.g. the NullReferenceException in SetArtworkThumbnail on
//     disable), and exceptions on SMTC callback threads can no longer take MusicBee down.
//   * The global keyboard/mouse hook (MouseKeyHook) was removed. It only fed a debounce
//     that was disabled (0 ms), so it did nothing except hook every key press.
//   * One reusable position timer instead of a new timer per play event.
//   * Artwork is written synchronously before DisplayUpdater.Update(), so the thumbnail
//     arrives together with the metadata.
//   * Problems are written to mb_MediaControl.log in MusicBee's persistent storage folder.

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Windows.Foundation;
using Windows.Media;
using Windows.Storage.Streams;

namespace MusicBeePlugin
{
    [ComImport]
    [Guid("ddb0472d-c911-4a1f-86d9-dc3d71a95f5a")]
    [InterfaceType(ComInterfaceType.InterfaceIsIInspectable)]
    internal interface ISystemMediaTransportControlsInterop
    {
        [return: MarshalAs(UnmanagedType.IInspectable)]
        object GetForWindow(IntPtr appWindow, [In] ref Guid riid);
    }

    public partial class Plugin
    {
        // IID of Windows.Media.ISystemMediaTransportControls
        private static readonly Guid IID_ISystemMediaTransportControls = new Guid("99FA3FF4-1742-42A6-902E-087D41F965EC");

        private MusicBeeApiInterface mbApiInterface;
        private readonly PluginInfo about = new PluginInfo();

        private readonly object sync = new object();
        private SystemMediaTransportControls smtc;
        private SystemMediaTransportControlsDisplayUpdater displayUpdater;
        private InMemoryRandomAccessStream artworkStream;
        private System.Threading.Timer positionTimer;
        private IntPtr registeredWindow = IntPtr.Zero;
        private volatile bool closing;
        private bool errorReported;
        private string logPath;

        public PluginInfo Initialise(IntPtr apiInterfacePtr)
        {
            mbApiInterface = new MusicBeeApiInterface();
            mbApiInterface.Initialise(apiInterfacePtr);
            about.PluginInfoVersion = PluginInfoVersion;
            about.Name = "Media Control";
            about.Description = "Enables MusicBee to interact with the Windows 10/11 media controls (SMTC). Patched build 1.1 (GetForWindow).";
            about.Author = "Steven Mayall, HenryPDT (patched)";
            about.TargetApplication = "";
            about.Type = PluginType.General;
            about.VersionMajor = 1;
            about.VersionMinor = 1;
            about.Revision = 0;
            about.MinInterfaceVersion = MinInterfaceVersion;
            about.MinApiRevision = MinApiRevision;
            about.ReceiveNotifications = (ReceiveNotificationFlags.PlayerEvents | ReceiveNotificationFlags.TagEvents);
            about.ConfigurationPanelHeight = 0;

            try
            {
                string dir = mbApiInterface.Setting_GetPersistentStoragePath();
                if (!string.IsNullOrEmpty(dir))
                    logPath = Path.Combine(dir, "mb_MediaControl.log");
            }
            catch
            {
                logPath = null;
            }

            return about;
        }

        public bool Configure(IntPtr panelHandle)
        {
            return false;
        }

        public void SaveSettings()
        {
        }

        public void Close(PluginCloseReason reason)
        {
            closing = true;
            TearDown();
        }

        public void Uninstall()
        {
        }

        // ------------------------------------------------------------------
        // MusicBee notifications
        // ------------------------------------------------------------------

        public void ReceiveNotification(string sourceFileUrl, NotificationType type)
        {
            if (closing)
                return;

            try
            {
                switch (type)
                {
                    case NotificationType.PluginStartup:
                    case NotificationType.MusicBeeStarted:
                        RunOnUiThreadAsync(() =>
                        {
                            if (smtc == null)
                                InitSmtc();
                        });
                        break;
                    case NotificationType.ApplicationWindowChanged:
                        RunOnUiThreadAsync(() =>
                        {
                            IntPtr hwnd = mbApiInterface.MB_GetWindowHandle();
                            if (hwnd != IntPtr.Zero && hwnd != registeredWindow)
                            {
                                Log("MusicBee window changed, re-registering media controls.");
                                TearDown();
                                InitSmtc();
                            }
                        });
                        break;
                    case NotificationType.PlayStateChanged:
                        UpdatePlayState();
                        break;
                    case NotificationType.TrackChanged:
                        UpdateDisplay();
                        UpdateTimeline();
                        break;
                    case NotificationType.NowPlayingArtworkReady:
                        UpdateDisplay();
                        break;
                    case NotificationType.TagsChanged:
                        if (sourceFileUrl != null && string.Equals(sourceFileUrl, mbApiInterface.NowPlaying_GetFileUrl(), StringComparison.OrdinalIgnoreCase))
                            UpdateDisplay();
                        break;
                    case NotificationType.PlayerShuffleChanged:
                        UpdateShuffle();
                        break;
                    case NotificationType.PlayerRepeatChanged:
                        UpdateRepeat();
                        break;
                }
            }
            catch (Exception ex)
            {
                Log("ReceiveNotification(" + type + ") failed: " + ex);
            }
        }

        // ------------------------------------------------------------------
        // SMTC setup / teardown
        // ------------------------------------------------------------------

        private void InitSmtc()
        {
            IntPtr hwnd = IntPtr.Zero;
            try { hwnd = mbApiInterface.MB_GetWindowHandle(); } catch (Exception ex) { Log("MB_GetWindowHandle failed: " + ex); }

            SystemMediaTransportControls controls = null;
            string method = null;

            if (hwnd != IntPtr.Zero)
            {
                try
                {
                    controls = GetForWindow(hwnd);
                    method = "GetForWindow";
                }
                catch (Exception ex)
                {
                    Log("GetForWindow failed: " + ex);
                }
            }
            else
            {
                Log("MusicBee window handle not available yet.");
            }

            if (controls == null)
            {
                // Old behaviour, kept only as a last resort (works on some older Windows 10 builds).
                try
                {
#pragma warning disable 618
                    controls = Windows.Media.Playback.BackgroundMediaPlayer.Current.SystemMediaTransportControls;
#pragma warning restore 618
                    method = "BackgroundMediaPlayer (fallback)";
                }
                catch (Exception ex)
                {
                    Log("BackgroundMediaPlayer fallback failed: " + ex);
                }
            }

            if (controls == null)
            {
                if (hwnd != IntPtr.Zero)
                    ReportError("Media Control could not register with the Windows media controls (SMTC).");
                return;
            }

            lock (sync)
            {
                smtc = controls;
                registeredWindow = hwnd;

                smtc.IsEnabled = true;
                smtc.IsPlayEnabled = true;
                smtc.IsPauseEnabled = true;
                smtc.IsStopEnabled = true;
                smtc.IsPreviousEnabled = true;
                smtc.IsNextEnabled = true;
                smtc.IsRewindEnabled = false;
                smtc.IsFastForwardEnabled = false;
                smtc.PlaybackStatus = MediaPlaybackStatus.Closed;
                TrySet(() => smtc.PlaybackRate = 1.0);

                smtc.ButtonPressed += Smtc_ButtonPressed;
                smtc.PlaybackPositionChangeRequested += Smtc_PlaybackPositionChangeRequested;
                smtc.ShuffleEnabledChangeRequested += Smtc_ShuffleEnabledChangeRequested;
                smtc.AutoRepeatModeChangeRequested += Smtc_AutoRepeatModeChangeRequested;

                displayUpdater = smtc.DisplayUpdater;

                if (positionTimer == null)
                    positionTimer = new System.Threading.Timer(_ => SafeRun(UpdateTimeline), null, Timeout.Infinite, Timeout.Infinite);
            }

            Log("Media controls registered via " + method + ".");

            SafeRun(UpdateDisplay);
            SafeRun(UpdateShuffle);
            SafeRun(UpdateRepeat);
            SafeRun(UpdatePlayState);
        }

        private static SystemMediaTransportControls GetForWindow(IntPtr hwnd)
        {
            object factory = WindowsRuntimeMarshal.GetActivationFactory(typeof(SystemMediaTransportControls));
            var interop = (ISystemMediaTransportControlsInterop)factory;
            Guid iid = IID_ISystemMediaTransportControls;
            object result = interop.GetForWindow(hwnd, ref iid);
            var controls = result as SystemMediaTransportControls;
            if (controls == null)
                throw new InvalidCastException("GetForWindow returned " + (result == null ? "null" : result.GetType().FullName));
            return controls;
        }

        // Runs the action on MusicBee's UI thread. Uses BeginInvoke (never Invoke) so a startup
        // notification delivered on a worker thread cannot deadlock against a busy UI thread.
        private void RunOnUiThreadAsync(Action action)
        {
            Control owner = null;
            try
            {
                IntPtr hwnd = mbApiInterface.MB_GetWindowHandle();
                if (hwnd != IntPtr.Zero)
                    owner = Control.FromHandle(hwnd);
            }
            catch
            {
                owner = null;
            }

            if (owner != null && owner.IsHandleCreated && owner.InvokeRequired)
            {
                try
                {
                    owner.BeginInvoke((Action)(() => SafeRun(action)));
                    return;
                }
                catch (Exception ex)
                {
                    Log("BeginInvoke failed, running inline: " + ex.Message);
                }
            }

            SafeRun(action);
        }

        private void TearDown()
        {
            lock (sync)
            {
                try { if (positionTimer != null) positionTimer.Dispose(); } catch { }
                positionTimer = null;

                if (smtc != null)
                {
                    try { smtc.ButtonPressed -= Smtc_ButtonPressed; } catch { }
                    try { smtc.PlaybackPositionChangeRequested -= Smtc_PlaybackPositionChangeRequested; } catch { }
                    try { smtc.ShuffleEnabledChangeRequested -= Smtc_ShuffleEnabledChangeRequested; } catch { }
                    try { smtc.AutoRepeatModeChangeRequested -= Smtc_AutoRepeatModeChangeRequested; } catch { }
                    try
                    {
                        if (displayUpdater != null)
                        {
                            displayUpdater.ClearAll();
                            displayUpdater.Update();
                        }
                    }
                    catch { }
                    try { smtc.PlaybackStatus = MediaPlaybackStatus.Closed; } catch { }
                    try { smtc.IsEnabled = false; } catch { }
                }

                try { if (artworkStream != null) artworkStream.Dispose(); } catch { }

                artworkStream = null;
                displayUpdater = null;
                smtc = null;
                registeredWindow = IntPtr.Zero;
            }
        }

        // ------------------------------------------------------------------
        // MusicBee -> SMTC
        // ------------------------------------------------------------------

        private void UpdatePlayState()
        {
            PlayState state = mbApiInterface.Player_GetPlayState();

            lock (sync)
            {
                if (smtc == null)
                    return;

                switch (state)
                {
                    case PlayState.Playing:
                        smtc.PlaybackStatus = MediaPlaybackStatus.Playing;
                        if (positionTimer != null) positionTimer.Change(0, 1000);
                        break;
                    case PlayState.Paused:
                        smtc.PlaybackStatus = MediaPlaybackStatus.Paused;
                        if (positionTimer != null) positionTimer.Change(Timeout.Infinite, Timeout.Infinite);
                        UpdateTimeline();
                        break;
                    case PlayState.Stopped:
                        smtc.PlaybackStatus = MediaPlaybackStatus.Stopped;
                        if (positionTimer != null) positionTimer.Change(Timeout.Infinite, Timeout.Infinite);
                        UpdateTimeline();
                        break;
                    case PlayState.Loading:
                        smtc.PlaybackStatus = MediaPlaybackStatus.Changing;
                        break;
                }
            }
        }

        private void UpdateDisplay()
        {
            string url = mbApiInterface.NowPlaying_GetFileUrl();

            string title = null, artist = null, albumArtist = null, album = null;
            uint trackNo = 0, trackCount = 0;
            byte[] artwork = null;

            if (!string.IsNullOrEmpty(url))
            {
                title = mbApiInterface.NowPlaying_GetFileTag(MetaDataType.TrackTitle);
                artist = mbApiInterface.NowPlaying_GetFileTag(MetaDataType.Artist);
                albumArtist = mbApiInterface.NowPlaying_GetFileTag(MetaDataType.AlbumArtist);
                album = mbApiInterface.NowPlaying_GetFileTag(MetaDataType.Album);
                uint.TryParse(mbApiInterface.NowPlaying_GetFileTag(MetaDataType.TrackNo), out trackNo);
                uint.TryParse(mbApiInterface.NowPlaying_GetFileTag(MetaDataType.TrackCount), out trackCount);

                if (string.IsNullOrEmpty(title))
                    title = url.Substring(url.LastIndexOfAny(new[] { '/', '\\' }) + 1);

                try
                {
                    PictureLocations locations;
                    string pictureUrl;
                    mbApiInterface.Library_GetArtworkEx(url, 0, true, out locations, out pictureUrl, out artwork);
                }
                catch (Exception ex)
                {
                    Log("Library_GetArtworkEx failed: " + ex.Message);
                    artwork = null;
                }
            }

            lock (sync)
            {
                if (displayUpdater == null)
                    return;

                displayUpdater.ClearAll();
                displayUpdater.Type = MediaPlaybackType.Music;

                if (!string.IsNullOrEmpty(url))
                {
                    MusicDisplayProperties music = displayUpdater.MusicProperties;
                    music.Title = title ?? string.Empty;
                    music.Artist = artist ?? string.Empty;
                    music.AlbumArtist = albumArtist ?? string.Empty;
                    TrySet(() => music.AlbumTitle = album ?? string.Empty);
                    if (trackNo > 0) TrySet(() => music.TrackNumber = trackNo);
                    if (trackCount > 0) TrySet(() => music.AlbumTrackCount = trackCount);
                }

                SetThumbnail(artwork);
                displayUpdater.Update();
            }
        }

        // must be called under lock(sync) with displayUpdater != null
        private void SetThumbnail(byte[] data)
        {
            InMemoryRandomAccessStream old = artworkStream;
            artworkStream = null;

            try
            {
                if (data == null || data.Length == 0)
                {
                    displayUpdater.Thumbnail = null;
                }
                else
                {
                    var stream = new InMemoryRandomAccessStream();
                    var writer = new DataWriter(stream);
                    writer.WriteBytes(data);
                    WaitFor(writer.StoreAsync());
                    WaitFor(writer.FlushAsync());
                    writer.DetachStream();
                    writer.Dispose();
                    stream.Seek(0);

                    artworkStream = stream;
                    displayUpdater.Thumbnail = RandomAccessStreamReference.CreateFromStream(stream);
                }
            }
            catch (Exception ex)
            {
                Log("Setting artwork failed: " + ex.Message);
                try { displayUpdater.Thumbnail = null; } catch { }
            }
            finally
            {
                try { if (old != null) old.Dispose(); } catch { }
            }
        }

        private static void WaitFor(IAsyncInfo operation)
        {
            var watch = Stopwatch.StartNew();
            while (operation.Status == AsyncStatus.Started)
            {
                if (watch.ElapsedMilliseconds > 3000)
                    throw new TimeoutException("WinRT stream operation timed out.");
                Thread.Sleep(1);
            }

            if (operation.Status != AsyncStatus.Completed)
            {
                Exception error = operation.ErrorCode;
                throw new IOException("WinRT stream operation ended with status " + operation.Status, error);
            }
        }

        private void UpdateTimeline()
        {
            int position = mbApiInterface.Player_GetPosition();
            int duration = mbApiInterface.NowPlaying_GetDuration();

            lock (sync)
            {
                if (smtc == null)
                    return;

                var timeline = new SystemMediaTransportControlsTimelineProperties();
                timeline.StartTime = TimeSpan.Zero;
                timeline.MinSeekTime = TimeSpan.Zero;
                timeline.Position = TimeSpan.FromMilliseconds(Math.Max(0, position));
                timeline.MaxSeekTime = TimeSpan.FromMilliseconds(Math.Max(0, duration));
                timeline.EndTime = TimeSpan.FromMilliseconds(Math.Max(0, duration));
                smtc.UpdateTimelineProperties(timeline);
            }
        }

        private void UpdateShuffle()
        {
            bool shuffle = mbApiInterface.Player_GetShuffle();
            lock (sync)
            {
                if (smtc != null)
                    smtc.ShuffleEnabled = shuffle;
            }
        }

        private void UpdateRepeat()
        {
            RepeatMode repeat = mbApiInterface.Player_GetRepeat();
            lock (sync)
            {
                if (smtc == null)
                    return;

                switch (repeat)
                {
                    case RepeatMode.One:
                        smtc.AutoRepeatMode = MediaPlaybackAutoRepeatMode.Track;
                        break;
                    case RepeatMode.All:
                        smtc.AutoRepeatMode = MediaPlaybackAutoRepeatMode.List;
                        break;
                    default:
                        smtc.AutoRepeatMode = MediaPlaybackAutoRepeatMode.None;
                        break;
                }
            }
        }

        // ------------------------------------------------------------------
        // SMTC -> MusicBee (called on WinRT threads: never let exceptions escape)
        // ------------------------------------------------------------------

        private void Smtc_ButtonPressed(SystemMediaTransportControls sender, SystemMediaTransportControlsButtonPressedEventArgs args)
        {
            SystemMediaTransportControlsButton button = args.Button;
            SafeRun(() =>
            {
                PlayState state = mbApiInterface.Player_GetPlayState();
                switch (button)
                {
                    case SystemMediaTransportControlsButton.Play:
                        if (state != PlayState.Playing)
                            mbApiInterface.Player_PlayPause();
                        break;
                    case SystemMediaTransportControlsButton.Pause:
                        if (state == PlayState.Playing)
                            mbApiInterface.Player_PlayPause();
                        break;
                    case SystemMediaTransportControlsButton.Stop:
                        mbApiInterface.Player_Stop();
                        break;
                    case SystemMediaTransportControlsButton.Next:
                        mbApiInterface.Player_PlayNextTrack();
                        break;
                    case SystemMediaTransportControlsButton.Previous:
                        mbApiInterface.Player_PlayPreviousTrack();
                        break;
                }
            });
        }

        private void Smtc_PlaybackPositionChangeRequested(SystemMediaTransportControls sender, PlaybackPositionChangeRequestedEventArgs args)
        {
            TimeSpan requested = args.RequestedPlaybackPosition;
            SafeRun(() =>
            {
                mbApiInterface.Player_SetPosition((int)Math.Max(0, requested.TotalMilliseconds));
                UpdateTimeline();
            });
        }

        private void Smtc_ShuffleEnabledChangeRequested(SystemMediaTransportControls sender, ShuffleEnabledChangeRequestedEventArgs args)
        {
            bool requested = args.RequestedShuffleEnabled;
            SafeRun(() =>
            {
                mbApiInterface.Player_SetShuffle(requested);
                UpdateShuffle();
            });
        }

        private void Smtc_AutoRepeatModeChangeRequested(SystemMediaTransportControls sender, AutoRepeatModeChangeRequestedEventArgs args)
        {
            MediaPlaybackAutoRepeatMode requested = args.RequestedAutoRepeatMode;
            SafeRun(() =>
            {
                switch (requested)
                {
                    case MediaPlaybackAutoRepeatMode.Track:
                        mbApiInterface.Player_SetRepeat(RepeatMode.One);
                        break;
                    case MediaPlaybackAutoRepeatMode.List:
                        mbApiInterface.Player_SetRepeat(RepeatMode.All);
                        break;
                    default:
                        mbApiInterface.Player_SetRepeat(RepeatMode.None);
                        break;
                }
                UpdateRepeat();
            });
        }

        // ------------------------------------------------------------------
        // helpers
        // ------------------------------------------------------------------

        private void SafeRun(Action action)
        {
            if (closing)
                return;
            try
            {
                action();
            }
            catch (Exception ex)
            {
                Log("Unhandled error: " + ex);
            }
        }

        private static void TrySet(Action action)
        {
            try { action(); } catch { }
        }

        private void ReportError(string message)
        {
            Log(message);
            if (errorReported)
                return;
            errorReported = true;

            string text = message + Environment.NewLine + Environment.NewLine +
                          "Details: " + (logPath ?? "(log path unavailable)");
            try
            {
                IntPtr hwnd = mbApiInterface.MB_GetWindowHandle();
                Control owner = hwnd != IntPtr.Zero ? Control.FromHandle(hwnd) : null;
                if (owner != null && owner.IsHandleCreated)
                    owner.BeginInvoke((Action)(() => MessageBox.Show(owner, text, "Media Control", MessageBoxButtons.OK, MessageBoxIcon.Warning)));
            }
            catch
            {
                // nothing else we can do; the log has the details
            }
        }

        private readonly object logLock = new object();

        private void Log(string message)
        {
            if (logPath == null)
                return;
            try
            {
                lock (logLock)
                {
                    var info = new FileInfo(logPath);
                    if (info.Exists && info.Length > 256 * 1024)
                        info.Delete();
                    File.AppendAllText(logPath, DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + message + Environment.NewLine, Encoding.UTF8);
                }
            }
            catch
            {
                // logging must never break the plugin
            }
        }
    }
}

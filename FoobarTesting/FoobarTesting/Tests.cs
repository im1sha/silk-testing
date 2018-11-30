using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using SilkTest.Ntf;
using SilkTest.Ntf.Win32;
using System.Threading;
using System.Runtime.InteropServices;

namespace FoobarTesting
{
    [SilkTestClass]
    public class Tests
    {
        private readonly Desktop _desktop = Agent.Desktop;

        readonly string[] folders = {
            "/Quick access/Google Drive",
            "__practice",
            "audios",
            "covers",
            "playlists"
        };
        readonly string app = "foobar2000 v1 4 1";
        readonly string grid = "{4B94B650-C2D8-40de-A0AD-E8FADF62D56C}";
        readonly string expectedProperty = "Properties (3 items) -  Whatever";
        readonly string playlist = "playlist";
        readonly string defaultPlaylist = "Default";
        readonly string expectedFileNames = "File names : Hollywood Undead - Whatever It Takes.mp3, " +
                "Thousand Foot Krutch - Scream.mp3, Three Days Grace - The High Road.mp3";

        [TestInitialize]
        public void Initialize()
        {
            BaseState baseState = new BaseState();
            baseState.Execute();
        }

        [TestMethod]
        public void OpenFile()
        {
            const string track = "Scream";
            const string trackFullName = "Thousand Foot Krutch - Scream";

            Window foobar2000V141 = _desktop.Window(app);
            foobar2000V141.SetActive();
            foobar2000V141.Control("ATL ToolbarWindow32").TextClick("File");
            foobar2000V141.MenuItem("Open").Select();

            Dialog open = foobar2000V141.Dialog("Open Dialog");
            open.Tree("Files of type").Select(folders[0]);
            open.TextField("File name").SetPosition(new TextPosition(0, 0));
            open.TextField("File name").SetText(folders[1]);
            open.SetActive();
            open.PushButton("Open").Select();
            open.TextField("File name").SetPosition(new TextPosition(0, 0));
            open.TextField("File name").SetText(folders[2]);
            open.SetActive();
            open.PushButton("Open").Select();
            open.TextField("File name").SetPosition(new TextPosition(0, 0));
            open.TextField("File name").SetText(trackFullName + ".mp3");
            open.PushButton("Open").Select();

            foobar2000V141.Control(grid).TextClick(track);

            foobar2000V141.Close();
        }

        private void LoadFiles(Window window, bool withRemoving = true)
        {

            if (withRemoving)
            {
                window.SetActive();
                window.Control(grid).TypeKeys("<Left Ctrl+a>");
                window.Control(grid).TypeKeys("<Delete>");
            }

            window.SetActive();
            window.Control("ATL ToolbarWindow32").TextClick("File");
            window.MenuItem("Add files").Select();

            Dialog addFiles = window.Dialog("Add Files");
            addFiles.Tree("Files of type").Select(folders[0]);
            addFiles.TextField("File name").SetPosition(new TextPosition(0, 0));
            addFiles.TextField("File name").SetText(folders[1]);
            addFiles.PushButton("Open").Select();
            addFiles.TextField("File name").SetText(folders[2]);
            addFiles.PushButton("Open").Select();
            addFiles.Control("DirectUIHWND").TypeKeys("<Left Ctrl+a>");

            addFiles.PushButton("Open").Select();

            if (withRemoving)
            {
                CheckProperties(window, expectedProperty);
            }
        }

        private void CheckProperties(Window window, string expectedProperty)
        {
            window.Control("ATL SysTabControl32").TextClick(defaultPlaylist);
            window.Control("ATL SysTabControl32").TextClick(defaultPlaylist, 1, ClickType.Right);
            window.MenuItem("Properties").Select();

            Dialog property = window.Dialog(expectedProperty);
            property.TabControl("TabControl").Select("Details");

            property.ListView("ListView").Select("File names");

            property.TypeKeys("<Left Ctrl+c>");
            Assert.AreEqual(expectedFileNames, Clipboard.GetText());
            property.PushButton("OK").Select();
        }

        [TestMethod]
        public void OpenFiles()
        {
            Window foobar2000V141 = _desktop.Window("foobar2000 v1 4 1");
            LoadFiles(foobar2000V141);
            foobar2000V141.Close();
        }

        [TestMethod]
        public void SaveAndLoadPlaylist()
        {
            Window foobar2000V141 = _desktop.Window("foobar2000 v1 4 1");
            LoadFiles(foobar2000V141);
            foobar2000V141.SetActive();
            foobar2000V141.Control("ATL ToolbarWindow32").TextClick("File");
            foobar2000V141.Control(grid).OpenContextMenu(new Point(184, 177));
            foobar2000V141.MenuItem("Save playlist").Select();

            Dialog savePlaylist = foobar2000V141.Dialog("Save Playlist");

            savePlaylist.Tree("Tree").Select(folders[0]);
            savePlaylist.TextField("TextField").SetPosition(new TextPosition(0, 0));
            savePlaylist.TextField("TextField").SetText(folders[1]);
            savePlaylist.PushButton("Save").Select();
            savePlaylist.TextField("TextField").SetPosition(new TextPosition(0, 0));
            savePlaylist.TextField("TextField").SetText(folders[4]);
            savePlaylist.PushButton("Save").Select();
            savePlaylist.TextField("TextField").SetPosition(new TextPosition(0, 0));
            savePlaylist.TextField("TextField").SetText(playlist);
            savePlaylist.TextField("TextField").TypeKeys("<Enter>");

            if (savePlaylist.GetChildren().Count != 0)
            {
                Dialog confirmSaveAs = savePlaylist.Dialog("Confirm Save As");
                confirmSaveAs.SetActive();
                confirmSaveAs.PushButton("Yes").Select();
            }

            foobar2000V141.SetActive();
            foobar2000V141.Control("ATL ToolbarWindow32").TextClick("File");
            foobar2000V141.MenuItem("Load playlist").Select();

            Dialog loadPlaylist = foobar2000V141.Dialog("Load Playlist");
            loadPlaylist.Control("DirectUIHWND").TextClick(playlist + ".fpl");
            loadPlaylist.SetActive();
            loadPlaylist.PushButton("Open").Select();
            foobar2000V141.Control("ATL SysTabControl32").TextClick(playlist, 1, ClickType.Right);
            foobar2000V141.Control("#32770").OpenContextMenu(new Point(76, 5));
            foobar2000V141.MenuItem("Properties").Select();

            Dialog property = foobar2000V141.Dialog(expectedProperty);
            property.PushButton("OK").Select();

            foobar2000V141.Control("ATL SysTabControl32").TextClick(playlist, 1, ClickType.Right);
            foobar2000V141.Control("#32770").OpenContextMenu(new Point(77, 11));
            foobar2000V141.MenuItem("Remove playlist").Select();

            foobar2000V141.Control("ATL SysTabControl32").TextClick(defaultPlaylist);
            foobar2000V141.Control("ATL SysTabControl32").TextClick(defaultPlaylist, 1, ClickType.Right);
            foobar2000V141.MenuItem("Properties").Select();

            property = foobar2000V141.Dialog(expectedProperty);
            property.PushButton("OK").Select();

            foobar2000V141.Close();
        }

        [TestMethod]
        public void Sort()
        {
            Window foobar2000V141 = _desktop.Window("foobar2000 v1 4 1");

            LoadFiles(foobar2000V141);

            const int totalTest = 2;
            const int checksInTest = 2;

            bool[] firstPositionIsLess = new bool[totalTest * checksInTest];
            int[] positionY = new int[2];

            for (int i = 0; i < checksInTest; i++)
            {
                foobar2000V141.Header("Header").Select("Artist/album", MouseButton.Left);
                foobar2000V141.Control(grid).TextClick("Hollywood Undead - Five");
                positionY[0] = Cursor.Position.Y;
                foobar2000V141.Control(grid).TextClick("Thousand Foot Krutch - Welcome To The Masquerade");
                positionY[1] = Cursor.Position.Y;
                firstPositionIsLess[i] = positionY[0] < positionY[1];
            }

            for (int i = 0; i < checksInTest; i++)
            {
                foobar2000V141.Header("Header").Select("Title / track artist", MouseButton.Left);
                foobar2000V141.Control(grid).TextClick("Whatever It Takes");
                positionY[0] = Cursor.Position.Y;
                foobar2000V141.Control(grid).TextClick("Scream");
                positionY[1] = Cursor.Position.Y;
                firstPositionIsLess[checksInTest + i] = positionY[0] < positionY[1];
            }

            Assert.AreNotEqual(firstPositionIsLess[0], firstPositionIsLess[1]);
            Assert.AreNotEqual(firstPositionIsLess[checksInTest + 0],
                firstPositionIsLess[checksInTest + 1]);

            foobar2000V141.Close();
        }

        [TestMethod]
        public void Play()
        {
            string[] expectedHeaders = {
                "Hollywood Undead -  Five #02  Whatever It Takes   foobar2000",
                "Thousand Foot Krutch -  Welcome",
                "Three Days Grace -  Transit Of"
            };

            Window foobar2000V141 = _desktop.Window(app);
            LoadFiles(foobar2000V141);
            foobar2000V141.Control("ATL ToolbarWindow32").TextClick("Playback");
            foobar2000V141.MenuItem("Stop").Select();

            foobar2000V141 = _desktop.Window(app);
            foobar2000V141.Control("ATL ToolbarWindow32").TextClick("Playback");
            foobar2000V141.MenuItem("Play").Select();

            for (int i = 0; i < expectedHeaders.Length; i++)
            {
                Window w = _desktop.Window(expectedHeaders[i]);
                w.Control("ATL ToolbarWindow32").TextClick("Playback");
                w.MenuItem("Next").Select();
            }

            foobar2000V141 = _desktop.Window(app);
            foobar2000V141.Close();
        }

        [TestMethod]
        public void Visualization()
        {
            Window foobar2000V141 = _desktop.Window(app);
            foobar2000V141.SetActive();
            foobar2000V141.Control("ATL ToolbarWindow32").TextClick("View");
            foobar2000V141.MenuItem("Oscilloscope").Select();

            Window oscilloscope = foobar2000V141.Window("Oscilloscope Window");
            oscilloscope.Close();
            foobar2000V141.Control("ATL ToolbarWindow32").TextClick("View");
            foobar2000V141.MenuItem("Peak Meter").Select();

            Window peakMeter = foobar2000V141.Window("Peak Meter Window");
            peakMeter.SetActive();
            peakMeter.Close();

            foobar2000V141.SetActive();
            foobar2000V141.Close();
        }

        [TestMethod]
        public void RemoveDuplicates()
        {
            Window foobar2000V141 = _desktop.Window(app);
            foobar2000V141.SetActive();

            LoadFiles(foobar2000V141);
            LoadFiles(foobar2000V141, false);

            foobar2000V141.Control(grid).SetFocus();
            foobar2000V141.Control("ATL ToolbarWindow32").TextClick("Edit");
            foobar2000V141.MenuItem("Remove duplicates").Select();

            CheckProperties(foobar2000V141, expectedProperty);

            foobar2000V141.Close();
        }

        [TestMethod]
        public void Rename()
        {
            Window foobar2000V141 = _desktop.Window("foobar2000 v1 4 1");

            LoadFiles(foobar2000V141);

            foobar2000V141.SetActive();

            string[] expectedNames = { "File name : Hollywood Undead - X.mp3" ,
                "File name : Hollywood Undead - Whatever It Takes.mp3" };
            string[] audio = { "Hollywood Undead - X" ,
                "Hollywood Undead - Whatever It Takes" };
            string album = "Hollywood Undead - Five";

            for (int i = 0; i < expectedNames.Length; i++)
            {
                foobar2000V141.Control(grid).TextClick(album);
                foobar2000V141.Control(grid).TextClick(album, 1, ClickType.Right);
                foobar2000V141.MenuItem("Rename to").Select();

                Dialog fileOperationsSetup = foobar2000V141.Dialog("File Operations Setup");
                fileOperationsSetup.SetActive();
                fileOperationsSetup.TextField("File name pattern").SetPosition(new TextPosition(0, 0));
                fileOperationsSetup.TextField("File name pattern").SetText(audio[i]);

                fileOperationsSetup.WaitForObject("Preview");
                fileOperationsSetup.PushButton("Run").Select();

                fileOperationsSetup.WaitForObject("Preview - nothing to do, destination matches source");
                fileOperationsSetup.PushButton("Close").Select();

                foobar2000V141.Control(grid).TextClick(album, 1, ClickType.Right);
                foobar2000V141.MenuItem("Properties").Select();

                Dialog audioDialog = foobar2000V141.Dialog("Properties -  Whatever It Takes");
                audioDialog.TabControl("TabControl").Select("Details");

                audioDialog.ListView("ListView").TextClick("File name", 1, ClickType.Right);
                audioDialog.MenuItem("Copy").Select();

                audioDialog.PushButton("OK").Select();

                Assert.AreEqual(expectedNames[i], Clipboard.GetText());
            }

            foobar2000V141.Close();
        }

        [TestMethod]
        public void Convert()
        {
            Window foobar2000V141 = _desktop.Window("foobar2000 v1 4 1");

            LoadFiles(foobar2000V141);

            foobar2000V141.Control(grid).TextClick("Hollywood Undead - Five");
            foobar2000V141.Control(grid).TextClick("Hollywood Undead - Five", 1, ClickType.Right);
            foobar2000V141.MenuItem("Quick convert").Select();

            Dialog quickConvert = foobar2000V141.Dialog("Quick Convert");
            quickConvert.ListView("ListView").Select("Opus");
            quickConvert.PushButton("Convert").Select();

            Dialog transcodeWarning = quickConvert.Dialog("Transcode warning");
            transcodeWarning.SetActive();
            transcodeWarning.PushButton("Yes").Select();

            Dialog saveAs = quickConvert.Dialog("Save As");
            saveAs.PushButton("Save").Select();
                
            Dialog converting = foobar2000V141.Dialog("Converting");
            converting.WaitForDisappearance(15000);

            Dialog converterOutput = foobar2000V141.Dialog("Converter Output");
            converterOutput.SetActive();
            converterOutput.Close();
            foobar2000V141.SetActive();
            foobar2000V141.Control("ATL ToolbarWindow32").TextClick("File");
            foobar2000V141.MenuItem("Open").Select();

            Dialog open = foobar2000V141.Dialog("Open Dialog");
            open.TextField("File name").SetPosition(new TextPosition(0, 0));
            open.TextField("File name").SetText("Hollywood Undead - Whatever It Takes.opus");
            open.PushButton("Open").Select();

            Window foobar2000NewWindow = _desktop.Window("Hollywood Undead -  Five #02  Whatever It Takes   foobar2000");

            foobar2000NewWindow.Control(grid).TextClick("Hollywood Undead - Five", 1, ClickType.Right);
            foobar2000NewWindow.Control(grid).OpenContextMenu(new Point(186, 28));
            foobar2000NewWindow.MenuItem("Delete file").Select();

            Dialog confirmFileRemoval = foobar2000NewWindow.Dialog("Confirm File Removal");
            confirmFileRemoval.PushButton("Yes").Select();

            foobar2000V141 = _desktop.Window("foobar2000 v1 4 1");
            foobar2000V141.WaitForObject("Playback stopped");

            foobar2000V141.Close();
        }

        [TestMethod]
        public void AttachFrontCoverFailed()
        {
            List<string> trackSizeStrings = new List<string>();

            Window foobar2000V141 = _desktop.Window("foobar2000 v1 4 1");
            foobar2000V141.SetActive();
            foobar2000V141.Control("ATL ToolbarWindow32").TextClick("File");
            foobar2000V141.MenuItem("Open").Select();

            Dialog open = foobar2000V141.Dialog("Open Dialog");
            open.Tree("Files of type").Select(folders[0]);
            open.SetActive();
            open.TextField("File name").SetPosition(new TextPosition(0, 0));
            open.TextField("File name").SetText(folders[1]);
            open.PushButton("Open").Select();
            open.TextField("File name").SetPosition(new TextPosition(0, 0));
            open.TextField("File name").SetText(folders[3]);
            open.PushButton("Open").Select();
            open.TextField("File name").SetText("Thousand Foot Krutch - Scream.mp3");
            open.PushButton("Open").Select();

            Window playingFoobar = _desktop.Window("Thousand Foot Krutch -  Welcome");
            playingFoobar.PushToolItem("PushToolItem2").Select();

            trackSizeStrings.Add(GetTrackSizeString(foobar2000V141));

            foobar2000V141 = _desktop.Window("foobar2000 v1 4 1");
            foobar2000V141.Control(grid).TextClick("Thousand Foot Krutch - Welcome To The Masquerade", 1, ClickType.Right);
            foobar2000V141.MenuItem("Manage attached pictures").Select();

            Dialog attachCoverDialog = foobar2000V141.Dialog("Attached Pictures  Thousand Foot Krutch - Scream mp3");
            attachCoverDialog.PushButton("Add").Select();
            attachCoverDialog.MenuItem("Front cover").Select();

            Dialog choosePictureFileToEmbed = attachCoverDialog.Dialog("Choose picture file to embed");
            choosePictureFileToEmbed.TextField("File name").SetPosition(new TextPosition(0, 0));
            choosePictureFileToEmbed.TextField("File name").SetText("Untitled design.jpg");
            choosePictureFileToEmbed.PushButton("Open").Select();
            attachCoverDialog.PushButton("OK").Select();

            trackSizeStrings.Add(GetTrackSizeString(foobar2000V141));

            foobar2000V141.Control(grid).TextClick("Thousand Foot Krutch - Welcome To The Masquerade", 1, ClickType.Right);
            foobar2000V141.MenuItem("Manage attached pictures").Select();

            Dialog attachedPicturesThousandFootKrutchScreamMp32 = foobar2000V141.Dialog("Attached Pictures  Thousand Foot Krutch - Scream mp3");
            attachedPicturesThousandFootKrutchScreamMp32.ListView("ListView").Select("Front cover : 4000x6000 JPEG");
            attachedPicturesThousandFootKrutchScreamMp32.PushButton("Remove").Select();
            attachedPicturesThousandFootKrutchScreamMp32.PushButton("OK").Select();

            trackSizeStrings.Add(GetTrackSizeString(foobar2000V141));

            Assert.AreEqual(trackSizeStrings[0], trackSizeStrings[1]);
            Assert.AreEqual(trackSizeStrings[0], trackSizeStrings[trackSizeStrings.Count - 1]);

            foobar2000V141.Close();
        }

        private string GetTrackSizeString(Window window, string track = "Thousand Foot Krutch - Welcome To The Masquerade", string dialogName = "Properties -  Scream")
        {
            window.Control(grid).TextClick(track, 1, ClickType.Right);
            window.MenuItem("Properties").Select();
            Dialog dialog = window.Dialog(dialogName);
            dialog.TabControl("TabControl").Select("Details");
            dialog.ListView("ListView").TextClick("File size", 1, ClickType.Right);
            dialog.MenuItem("Copy").Select();
            dialog.PushButton("OK").Select();
            return Clipboard.GetText();
        }
    }
}


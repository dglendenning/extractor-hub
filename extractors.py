"""Compilation of LinkIt data extractors in one UI window."""

import wx
import extract_benchmark
import extract_parcc
import usage_report
import benchmark_status

__version__ = None  # set in main()


class ExtractFrame(wx.Frame):
    """Cross-platform window UI for Malcolm's extract utilities."""

    def __init__(self, *args, **kw):
        """Call parent constructor and build UI elements."""
        # ensure the parent's __init__ is called
        super(ExtractFrame, self).__init__(*args, **kw)
        # create a panel in the frame
        pnl = wx.Panel(self)
        # added by mal til they figure out how to access Panels
        self.panel = pnl
        # create a menu bar
        self.makeMenuBar()

        # and a status bar
        self.CreateStatusBar()
        self.SetStatusText("Status: Idle")

    def makeMenuBar(self):
        """Build menu bar and bind methods to each item."""
        # Make a file menu with Hello and Exit items
        fileMenu = wx.Menu()
        fileMenu.AppendSeparator()
        # When using a stock ID we don't need to specify the menu item's
        # label
        exitItem = fileMenu.Append(wx.ID_EXIT)
        # Now a help menu for the about item
        helpMenu = wx.Menu()
        aboutItem = helpMenu.Append(wx.ID_ABOUT)
        self.extractMenu = wx.Menu()

        self.AddExtract(
            "Extract Benchmark", self.BenchmarkExtract,
            help="Extract data for Benchmark Navigator", key="B")
        self.AddExtract(
            "Extract PARCC", self.PARCCExtract,
            help="Extract data for PARCC Report", key="P")
        self.AddExtract(
            "Weekly Usage Report", self.UsageReport,
            help="Create report of this past week's usage data (FRI - THU)",
            key="U")
        self.AddExtract(
            "Get Benchmark Status", self.BenchmarkStatus,
            help="Find which students are/are not finished with a Benchmark.",
            key="S")
        # Make the menu bar and add the three menus to it. The '&' defines
        # that the next letter is the "mnemonic" for the menu item. On the
        # platforms that support it those letters are underlined and can be
        # triggered from the keyboard.
        menuBar = wx.MenuBar()
        menuBar.Append(fileMenu, "&File")
        menuBar.Append(helpMenu, "&Help")
        menuBar.Append(self.extractMenu, "&Extract")

        # Gives the menu bar to the frame (can be edited/appended afterward)
        self.SetMenuBar(menuBar)

        # Finally, associate a handler function with the EVT_MENU event for
        # each of the menu items. That means that when that menu item is
        # activated then the associated handler function will be called.
        # NOTE: Not necessary if using self.AddExtract()
        self.Bind(wx.EVT_MENU, self.OnExit,  exitItem)
        self.Bind(wx.EVT_MENU, self.OnAbout, aboutItem)

    def OnExit(self, event):
        """Close the frame, terminating the application."""
        self.Close(True)

    def BenchmarkExtract(self, event):
        """Extract Benchmark for Navigator Report."""
        districtID = self.getDistrictID()
        if districtID is None:
            return False

        # Code goes here - we have a valid districtID by this point.
        self.SetStatusText("Status: Extracting...")
        m = extract_benchmark.extract(districtID)
        self.SetStatusText("Status: Idle")
        wx.MessageBox(m)
        return True

    def BenchmarkStatus(self, event):
        """Get Benchmark completion status."""
        districtID = self.getDistrictID()
        if districtID is None:
            return False
        form = self.getForm()
        if form not in ["A", "B", "C"]:
            return False
        # Code goes here - we have a valid districtID by this point.
        self.SetStatusText("Status: Extracting...")
        m = benchmark_status.main(districtID, form)
        self.SetStatusText("Status: Idle")
        wx.MessageBox(m)
        return True

    def PARCCExtract(self, event):
        """Extract PARCC for Navigator Report."""
        districtID = self.getDistrictID()
        if districtID is None:
            return False

        # Code goes here - we have a valid districtID by this point.
        self.SetStatusText("Status: Extracting...")
        m = extract_parcc.extract(str(districtID))
        self.SetStatusText("Status: Idle")
        wx.MessageBox(m)
        return True

    def UsageReport(self, event):
        """Create Weekly Usage Report."""
        self.SetStatusText("Status: Extracting...")
        if usage_report.create_report():
            wx.MessageBox("Report Created!")
        else:
            wx.MessageBox("Report Failed.")
        self.SetStatusText("Status: Idle")

    def getDistrictID(self):
        """Query user for DistrictID and return it as a string."""
        dialog = wx.TextEntryDialog(
            self, "Please enter DistrictID", "DistrictID")

        while True:
            if dialog.ShowModal() == wx.ID_CANCEL:
                break
            s = dialog.GetValue()

            try:
                s = int(s)
            except ValueError as ex:
                wx.MessageBox("Value Error: " + str(ex) + "\n"
                              + "Try again. "
                              + "Click cancel on next window to quit.")
                dialog.SetValue("")
            else:
                return s

    def getForm(self):
        """Get A, B or C from user."""
        dialog = wx.SingleChoiceDialog(self, "Select a Form", "Form",
                                       choices=["A", "B", "C"])
        while True:
            if dialog.ShowModal() == wx.ID_CANCEL:
                break
            choice = dialog.GetSelection()
            return ["A", "B", "C"][choice]

    def AddExtract(self, name, event, help="", key=None):
        """Add an item to the extract menu on the menu bar.

        Keyword arguments:
        name -- appears on the label in the Extract menu
        event -- method of this class to run when clicked
        help -- Help string displayed on status bar when hovered over
        key -- Hotkey to run event
        """
        if name[0] != "&":
            name = "&" + name

        if key is not None:
            name = name + "\t" + key

        item = self.extractMenu.Append(-1, name, help)
        self.Bind(wx.EVT_MENU, event, item)

    def OnAbout(self, event):
        """Display an About Dialog with the version number."""
        wx.MessageBox("Ask Malcolm!",
                      "Extractor Hub v{}".format(__version__),
                      wx.OK | wx.ICON_INFORMATION)


def main(version='0.0.0'):
    """Launch an ExtractFrame."""
    global __version__
    __version__ = version
    app = wx.App()
    frm = ExtractFrame(None, title='Extractor Hub')
    frm.Show()
    app.MainLoop()

import os
import wx
import wx.lib.agw.aui as aui
import win32com.client
from PyPDF2 import PdfMerger
from pathlib import Path
from cryptography.fernet import Fernet


class PDFMergerApp(wx.App):
    def OnInit(self):
        self.frame = MainWindow(None, title="PDF Merger")
        self.SetTopWindow(self.frame)
        self.frame.Show()
        return True


class MainWindow(wx.Frame):
    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)

        self.pdfs = []

        self.manager = aui.AuiManager(self)

        self.panel = wx.Panel(self)
        sizer = wx.BoxSizer(wx.VERTICAL)

        self.pdf_listbox = wx.ListBox(self.panel, style=wx.LB_EXTENDED)
        sizer.Add(self.pdf_listbox, 1, wx.EXPAND | wx.ALL, 10)

        select_button = wx.Button(self.panel, label="Select Files")
        select_button.Bind(wx.EVT_BUTTON, self.on_select)
        sizer.Add(select_button, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        merge_button = wx.Button(self.panel, label="Merge PDFs")
        merge_button.Bind(wx.EVT_BUTTON, self.on_merge)
        sizer.Add(merge_button, 0, wx.ALL | wx.ALIGN_CENTER, 10)

        self.panel.SetSizerAndFit(sizer)

        self.manager.AddPane(
            self.panel,
            aui.AuiPaneInfo().CenterPane().Name("centerpane")
        )

        self.manager.Update()

    def on_select(self, event):
        file_dialog = wx.FileDialog(
            self,
            message="Select Files",
            defaultDir=os.getcwd(),
            defaultFile="",
            wildcard="Supported files (*.pdf;*.docx;*.ppt;*.xls)|*.pdf;*.docx;*.ppt;*.xls",
            style=wx.FD_OPEN | wx.FD_FILE_MUST_EXIST | wx.FD_MULTIPLE
        )

        if file_dialog.ShowModal() == wx.ID_OK:
            selected_files = file_dialog.GetPaths()
            for file in selected_files:
                self.pdfs.append(file)
                self.pdf_listbox.Append(Path(file).name)

        file_dialog.Destroy()

    def on_merge(self, event):
        if len(self.pdfs) == 0:
            wx.MessageBox("No files selected.", "Error", wx.OK | wx.ICON_ERROR)
            return

        merger = PdfMerger()
        for file_path in self.pdfs:
            extension = Path(file_path).suffix.lower()

            if extension == ".pdf":
                merger.append(file_path)
            else:
                pdf_path = self.convert_to_pdf(file_path)
                if pdf_path:
                    merger.append(pdf_path)

        dlg = wx.FileDialog(
            self,
            message="Save merged PDF file",
            defaultDir=os.getcwd(),
            defaultFile="merged.pdf",
            wildcard="PDF files (*.pdf)|*.pdf",
            style=wx.FD_SAVE | wx.FD_OVERWRITE_PROMPT
        )

        if dlg.ShowModal() == wx.ID_OK:
            output_filename = dlg.GetPath()

            # Encrypt the merged PDF file
            encrypted_filename = self.encrypt_pdf(merger, output_filename)
            merger.close()

            wx.MessageBox(f"PDFs merged and encrypted. Encrypted file saved as {encrypted_filename}", "Success",
                          wx.OK | wx.ICON_INFORMATION)

            # Open the encrypted PDF file
            os.startfile(encrypted_filename)

        dlg.Destroy()

    @staticmethod
    def convert_to_pdf(file_path):
        # Conversion logic as before
        pass

    @staticmethod
    def encrypt_pdf(merger, output_filename):
        # Save the merged PDF file
        merged_filename = "merged.pdf"
        merger.write(merged_filename)

        # Generate encryption key
        key = Fernet.generate_key()
        fernet = Fernet(key)

        # Read the merged PDF file
        with open(merged_filename, "rb") as file:
            pdf_data = file.read()

        # Encrypt the PDF data
        encrypted_data = fernet.encrypt(pdf_data)

        # Save the encrypted PDF file
        encrypted_filename = output_filename.replace(".pdf", "_encrypted.pdf")
        with open(encrypted_filename, "wb") as file:
            file.write(encrypted_data)

        # Clean up temporary files
        os.remove(merged_filename)

        return encrypted_filename


if __name__ == "__main__":
    app = PDFMergerApp()
    app.MainLoop()

    import wx

    app = wx.App()

    dlg = wx.MessageDialog(None, "Hello, World!", "Dialog Title", wx.OK | wx.ICON_INFORMATION)
    dlg.ShowModal()
    dlg.Destroy()

    app.MainLoop()


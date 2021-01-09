from PyPDF2 import PdfFileMerger, PdfFileReader, PdfFileWriter
import os
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import *

class pdf_manipulate(QWidget):
    def __init__(self):
        super().__init__()
        self.title = "PDF merge"
        self.cornerX = 130
        self.cornerY = 50
        self.width = 500
        self.height = 500
        self.max_num_of_merge_files = 30
        self.merge_files_label = QLabel("Files to merge", self)
        self.split_file_label = QLabel("File to split", self)
        self.split_pages_label = QLabel("Pages to extract", self)
        self.split_pages_label.setAlignment(Qt.AlignCenter)
        self.line_page_numbers = []
        self.line_list_merger = []
        self.line_list_splitter = []
        self.button_list_merger = []
        self.button_list_splitter = []
        self.hbox_list_merger = []
        self.hbox_list_splitter = []
        self.num_of_files_merger = 0
        self.num_of_files_splitter = 0
        self.file_name_dict_merger = {}
        self.file_name_dict_splitter = {}
        self.input_error_flag = False
        self.tabs = QTabWidget()
        self.merge_tab = QWidget()
        self.split_tab = QWidget()
        self.initUI()

    def initUI(self):
        # Set basic UI settings and layout
        self.tabs.addTab(self.merge_tab, "Merge")
        self.tabs.addTab(self.split_tab, "Split")
        self.merge_tab.grid = QGridLayout()
        self.split_tab.grid = QGridLayout()
        self.grid = QGridLayout()
        self.setLayout(self.grid)
        self.merge_tab.setLayout(self.merge_tab.grid)
        self.split_tab.setLayout(self.split_tab.grid)
        self.setGeometry(self.cornerX, self.cornerY, self.width, self.height)
        self.setWindowTitle(self.title)
        self.vertical_spacer_merger = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)
        self.vertical_spacer_splitter = QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding)

        # Initialize scroll area, text boxes and buttons
        self.line_list_merger.append(QLineEdit(self))
        self.line_list_merger[0].setReadOnly(True)
        self.button_list_merger.append(QPushButton("Select file"))
        self.hbox_list_merger.append(QHBoxLayout())
        self.hbox_list_merger[self.num_of_files_merger].addWidget(self.line_list_merger[self.num_of_files_merger])
        self.hbox_list_merger[self.num_of_files_merger].addWidget(self.button_list_merger[self.num_of_files_merger])
        self.merge_scroll_area = QScrollArea(self)
        self.merge_scroll_area.setWidgetResizable(True)
        self.merge_scroll_area_widget_contents = QWidget()
        self.merge_scroll_area.setWidget(self.merge_scroll_area_widget_contents)
        self.vbox_merger = QVBoxLayout(self.merge_scroll_area_widget_contents)
        self.vbox_merger.addLayout(self.hbox_list_merger[self.num_of_files_merger])
        self.vbox_merger.addItem(self.vertical_spacer_merger)
        self.add_files_button = QPushButton("New file")
        self.remove_files_button = QPushButton("Remove file")
        self.merge_button = QPushButton("Merge")

        self.line_list_splitter.append(QLineEdit(self))
        self.line_page_numbers.append(QLineEdit(self))
        self.line_list_splitter[0].setReadOnly(True)
        self.button_list_splitter.append(QPushButton("Select file"))
        self.hbox_list_splitter.append(QHBoxLayout())
        self.hbox_list_splitter[self.num_of_files_splitter].addWidget(self.line_list_splitter[self.num_of_files_splitter])
        self.hbox_list_splitter[self.num_of_files_splitter].addWidget(self.line_page_numbers[self.num_of_files_splitter])
        self.hbox_list_splitter[self.num_of_files_splitter].addWidget(self.button_list_splitter[self.num_of_files_splitter])
        self.split_scroll_area = QScrollArea(self)
        self.split_scroll_area.setWidgetResizable(True)
        self.split_scroll_area_widget_contents = QWidget()
        self.split_scroll_area.setWidget(self.split_scroll_area_widget_contents)
        self.vbox_splitter = QVBoxLayout(self.split_scroll_area_widget_contents)
        self.vbox_splitter.addLayout(self.hbox_list_splitter[self.num_of_files_splitter])
        self.vbox_splitter.addItem(self.vertical_spacer_splitter)
        self.split_button = QPushButton("Split")

        # Add widgets to layout
        self.grid.addWidget(self.tabs,0,0)

        self.merge_tab.grid.addWidget(self.merge_files_label,0,0)
        self.merge_tab.grid.addWidget(self.merge_scroll_area,1,0)
        self.merge_tab.grid.addWidget(self.add_files_button,2,0)
        self.merge_tab.grid.addWidget(self.remove_files_button,3,0)
        self.merge_tab.grid.addWidget(self.merge_button,4,0)

        self.split_tab.grid.addWidget(self.split_file_label,0,0)
        self.split_tab.grid.addWidget(self.split_pages_label,0,0)
        self.split_tab.grid.addWidget(self.split_scroll_area,1,0)
        self.split_tab.grid.addWidget(self.split_button,2,0)

        # Connect user actions to functions
        self.add_files_button.clicked.connect(self.add_new_file)
        self.remove_files_button.clicked.connect(self.remove_file)
        self.merge_button.clicked.connect(self.merge_files)
        self.split_button.clicked.connect(self.split_file)
        idx = len(self.button_list_merger) - 1
        self.button_list_merger[self.num_of_files_merger].clicked.connect(lambda: self.select_input_file_merger(idx))
        self.button_list_splitter[self.num_of_files_splitter].clicked.connect(lambda: self.select_input_file_splitter(idx))
        self.show()

    # Add new input file for merging
    def add_new_file(self):
        if self.num_of_files_merger > self.max_num_of_merge_files - 2:
            message = "Cannot merge more than {} files.".format(self.max_num_of_merge_files)
            QMessageBox.about(self, "Error", message)
            return
        self.num_of_files_merger += 1
        self.line_list_merger.append(QLineEdit(self))
        self.button_list_merger.append(QPushButton("Select file"))
        self.hbox_list_merger.append(QHBoxLayout())
        self.hbox_list_merger[self.num_of_files_merger].addWidget(self.line_list_merger[self.num_of_files_merger])
        self.hbox_list_merger[self.num_of_files_merger].addWidget(self.button_list_merger[self.num_of_files_merger])
        self.vbox_merger.insertLayout(self.num_of_files_merger, self.hbox_list_merger[self.num_of_files_merger])
        self.line_list_merger[self.num_of_files_merger].setReadOnly(True);
        idx = len(self.button_list_merger) - 1
        self.button_list_merger[self.num_of_files_merger].clicked.connect(lambda: self.select_input_file_merger(idx))

    # Remove the latest file to be merged
    def remove_file(self):
        if (self.num_of_files_merger >= 1):
            self.line_list_merger.pop().setParent(None)
            self.button_list_merger.pop().setParent(None)
            self.hbox_list_merger.pop().setParent(None)
            self.num_of_files_merger -= 1

    def select_input_file_merger(self, idx):
        options = QFileDialog.Options()
        file_name_full, _ = QFileDialog.getOpenFileName(self, "Select file", "", "PDF files (*.pdf)", options=options)
        file_name_list = file_name_full.split("/")
        file_name = file_name_list[len(file_name_list) - 1]
        self.line_list_merger[idx].setText(file_name)
        self.file_name_dict_merger[idx] = file_name_full

    def select_input_file_splitter(self, idx):
        options = QFileDialog.Options()
        file_name_full, _ = QFileDialog.getOpenFileName(self, "Select file", "", "PDF files (*.pdf)", options=options)
        file_name_list = file_name_full.split("/")
        file_name = file_name_list[len(file_name_list) - 1]
        self.line_list_splitter[idx].setText(file_name)
        self.file_name_dict_splitter[idx] = file_name_full

    def merge_files(self):
        merger = PdfFileMerger()
        for i in range(self.num_of_files_merger + 1):
            if (not self.line_list_merger[i].text() or self.line_list_merger[i].text() == "INVALID FILE"):
                self.line_list_merger[i].setText("INVALID FILE")
                self.input_error_flag = True
        if (self.input_error_flag):
            QMessageBox.about(self, "Error", "An input file cannot be empty")
            self.input_error_flag = False
            return
        options = QFileDialog.Options()
        save_file_name = QFileDialog.getSaveFileName(self, "Select destination file", "", "PDF files (*.pdf)", options=options)
        if (save_file_name[0] == ""):
            QMessageBox.about(self, "Error", "Save file name cannot be empty")
            return
        for i in range(self.num_of_files_merger + 1):
            merger.append(PdfFileReader(self.file_name_dict_merger[i], 'rb'))
        merger.write(save_file_name[0])
        merger.close()
        QMessageBox.about(self, "Done", "Merge succeeded")

    def split_file(self):
        for i in range(self.num_of_files_splitter + 1):
            if (not self.line_list_splitter[i].text() or self.line_list_splitter[i].text() == "INVALID FILE"):
                self.line_list_splitter[i].setText("INVALID FILE")
                self.input_error_flag = True
        if (self.input_error_flag):
            QMessageBox.about(self, "Error", "An input file cannot be empty")
            self.input_error_flag = False
            return
        reader = PdfFileReader(self.file_name_dict_splitter[0], strict=False)
        pages_str = self.line_page_numbers[0].text()
        save_dir_name = QFileDialog.getExistingDirectory(self, "Select destination directory")
        pages_full = pages_str.split(",")
        for pages in pages_full:
            if not ":" in pages:
                if (not pages.isnumeric()):
                    QMessageBox.about(self, "Error", "Invalid pages to extract")
                    return
                writer = PdfFileWriter()
                writer.addPage(reader.getPage(int(pages) - 1))
                file_name = str(save_dir_name + "/" + "page_" + pages + ".pdf")
                with open(file_name, 'wb') as outfile:
                    writer.write(outfile)
            elif ":" in pages:
                writer = PdfFileWriter()
                page = pages.split(":")
                if ((not(page[0].isnumeric())) or (not(page[1].isnumeric()))):
                    QMessageBox.about(self, "Error", "Invalid pages to extract")
                    return
                first_page = int(page[0]) - 1
                last_page = int(page[1]) - 1
                i = first_page
                while (i <= last_page):
                    writer.addPage(reader.getPage(int(i)))
                    i += 1
                file_name = str(save_dir_name + "/" + "pages_" + str(page[0]) + "_" + str(page[1]) + ".pdf")
                with open(file_name, 'wb') as outfile:
                    writer.write(outfile)
        QMessageBox.about(self, "Done", "Split succeeded")

app = QApplication(sys.argv)
manipulate = pdf_manipulate()

sys.exit(app.exec())

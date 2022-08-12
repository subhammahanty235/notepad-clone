from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtPrintSupport import *
from PyQt5.QtCore import *
import docx2txt
import sys


class BsWord(QMainWindow):
    def __init__(self):
        super(BsWord , self).__init__()
        self.editor = QTextEdit()
        self.editor.setFontPointSize(20)
        # self.editor.setFontFamily("Helvetica")
        self.setCentralWidget(self.editor)
        self.font_size_box = QSpinBox()
        self.showMaximized()
        self.setWindowTitle('Chilz Editor')
        self.create_tool_bar()
        self.create_menu_bar()

        # stores path
        self.path = ''
    def create_menu_bar(self):
        menubar = QMenuBar()

        # app_icon = menubar.addMenu(QIcon("doc_icon.png"),"icon")

        file_menu = QMenu('File' , self)
        menubar.addMenu(file_menu)

        save_as_pdf_action = QAction('Save As PDF' , self)
        save_as_pdf_action.triggered.connect(self.save_as_pdf_func)
        file_menu.addAction(save_as_pdf_action)


        save_action = QAction('Save' ,self)
        save_action.triggered.connect(self.file_save)
        file_menu.addAction(save_action)

        open_action = QAction('Open' ,self)
        open_action.triggered.connect(self.open_file)
        file_menu.addAction(open_action)

        rename_action = QAction('Rename' ,self)
        rename_action.triggered.connect(self.rename_file)
        file_menu.addAction(rename_action)

        # Edit manu
        edit_menu = QMenu('Edit' , self)
        menubar.addMenu(edit_menu)
        # paste operation button
        paste_action = QAction('Paste' ,self)
        paste_action.triggered.connect(self.editor.paste)
        edit_menu.addAction(paste_action)

        # Clear operation button
        clear_menu = QAction('Clear' , self)
        clear_menu.triggered.connect(self.editor.clear)
        edit_menu.addAction(clear_menu)

        # selectall operation button
        selectAll_menu = QAction('Select All' , self)
        selectAll_menu.triggered.connect(self.editor.selectAll)
        edit_menu.addAction(selectAll_menu)

        # view mwnu
        view_menu = QMenu('View' , self)
        menubar.addMenu(view_menu)
        # set editor to fullscreen
        fullscreen_action = QAction('FullScreen' ,self)
        fullscreen_action.triggered.connect(lambda:self.showFullScreen())
        view_menu.addAction(fullscreen_action)

        # set Normal View
        normscr_action = QAction('Normal View', self)
        normscr_action.triggered.connect(lambda : self.showNormal())
        view_menu.addAction(normscr_action)

        # minimize the editor
        minscr_action = QAction('Minimize', self)
        minscr_action.triggered.connect(lambda : self.showMinimized())
        view_menu.addAction(minscr_action)


        self.setMenuBar(menubar)
    def create_tool_bar(self):
        toolbar = QToolBar()
        # undo operation
        undo_action = QAction(QIcon('undo.png'),'undo' , self)
        undo_action.triggered.connect(self.editor.undo)
        toolbar.addAction(undo_action)
        # redo operation
        redo_action = QAction(QIcon('redo.png'),'redo' , self)
        redo_action.triggered.connect(self.editor.redo)
        toolbar.addAction(redo_action)
        # add a separator
        toolbar.addSeparator()
        toolbar.addSeparator()
        # cut operation
        cut_action = QAction(QIcon('cut.png'),'cut' , self)
        cut_action.triggered.connect(self.editor.cut)
        toolbar.addAction(cut_action)
        # copy operation
        copy_action = QAction(QIcon('copy.png'),'copy' , self)
        copy_action.triggered.connect(self.editor.copy)
        toolbar.addAction(copy_action)
        # paste operation
        paste_action = QAction(QIcon('paste.png'),'paste' , self)
        paste_action.triggered.connect(self.editor.paste)
        toolbar.addAction(paste_action)
         # add a separator
        toolbar.addSeparator()
        toolbar.addSeparator()
        # font size change operation
        self.font_size_box.setValue(20)
        self.font_size_box.valueChanged.connect(self.setFontSize)
        toolbar.addWidget(self.font_size_box)

        # font change operation
        toolbar.addSeparator()
        self.font_combo = QComboBox(self)
        self.font_combo.addItems(["Courier Std", "Hellentic Typewriter Regular", "Helvetica", "Arial", "SansSerif", "Helvetica", "Times", "Monospace"])
        self.font_combo.activated.connect(self.set_font)      # connect with function
        toolbar.addWidget(self.font_combo) 
        toolbar.addSeparator()
        toolbar.addSeparator()
        
        # bold
        bold_action = QAction(QIcon("bold.png"), 'Bold', self)
        bold_action.triggered.connect(self.bold_text)
        toolbar.addAction(bold_action)

        # italic
        italic_action = QAction(QIcon('italic.png') , 'Italic' , self);
        italic_action.triggered.connect(self.italic_text)
        toolbar.addAction(italic_action)

        # underline
        underline_action = QAction(QIcon("underline.png"), 'Underline', self)
        underline_action.triggered.connect(self.underline_text)
        toolbar.addAction(underline_action)
        toolbar.addSeparator()


        self.addToolBar(toolbar)
    
    
    def setFontSize(self):
        value = self.font_size_box.value()
        self.editor.setFontPointSize(value)

    def save_as_pdf_func(self):
        file_path , _ = QFileDialog.getSaveFileName(self , 'Export PDF', None , 'PDF Files (*.pdf)')
        printer = QPrinter(QPrinter.HighResolution)
        printer.setOutputFormat(QPrinter.PdfFormat)
        printer.setOutputFileName(file_path)
        self.editor.document().print_(printer)

    def open_file(self):
        self.path , _ = QFileDialog.getOpenFileName(self, "open file" ,"", "Text documents (*.text);Text documents (*.txt); All files (*.*)")
        try:
            text = docx2txt.process(self.path)
        except Exception as e :
            print(e)
        
        else:
            self.editor.setText(text)
            self.update_title();
    def rename_file(self):
        self.setWindowTitle(self.title + ' ' + self.path)

    def file_saveas(self):
        self.path, _ = QFileDialog.getSaveFileName(self, "Save file", "", "text documents (*.text);Text documents (*.txt);All files (*.*)")

        if self.path == '':
            return   # If dialog is cancelled, will return ''

        text = self.editor.toPlainText()

        try:
            with open(path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)
    def file_save(self):
        print(self.path)
        if self.path == '':
            # If we do not have a path, we need to use Save As.
            self.file_saveas()

        text = self.editor.toPlainText()

        try:
            with open(self.path, 'w') as f:
                f.write(text)
                self.update_title()
        except Exception as e:
            print(e)

    def set_font(self):
        font = self.font_combo.currentText()
        self.editor.setCurrentFont(QFont(font))

    def bold_text(self):
        if self.editor.fontWeight()!= QFont.Bold:
            self.editor.setFontWeight(QFont.Bold)
            return
        self.editor.setFontWeight(QFont.Normal)
    def underline_text(self):
        state = self.editor.fontUnderline()
        print(state)
        self.editor.setFontUnderline(not(state))
    def italic_text(self):
        state = self.editor.fontItalic()
        self.editor.setFontItalic(not(state))
app = QApplication(sys.argv)
window = BsWord()
window.show()
sys.exit(app.exec_())
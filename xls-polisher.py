# Copyright (c) 2014 Giulio De Pasquale
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.

# "Universal v0.3" Icon Theme by Freepik is licensed under CC BY 3.0

__author__ = 'Giulio De Pasquale'

import sys
from tempfile import TemporaryFile
from collections import defaultdict
from PyQt4 import QtGui, uic, QtCore

import xlrd
from xlwt import Workbook


main_window_ui = uic.loadUiType("./ui/qt-main.ui")[0]
dialog_window_ui = uic.loadUiType("./ui/qt-about.ui")[0]
filter_window_ui = uic.loadUiType("./ui/qt-filter.ui")[0]
remove_column_dialog_ui = uic.loadUiType("./ui/qt-remove-column.ui")[0]
tab_widget_ui = uic.loadUiType("./ui/qt-tabwidget.ui")[0]
cross_icon_path = "./ui/resources/icons/cross18.png"


def filenamefromopenfiledialog(parent=None):
    return QtGui.QFileDialog.getOpenFileName(parent, "XLS Polisher - Open Excel File",
                                             "/home",
                                             "Excel Files (*.xls *.xlsx)")


def filenamefromsavefiledialog(parent=None):
    return QtGui.QFileDialog.getSaveFileName(parent, "XLS Polisher - Save Polished File",
                                             "/home",
                                             "Excel Files (*.xls *.xlsx)")


class FilterDetails():
    def __init__(self, colname, strict, show, string):
        self.colName = colname
        self.strict = strict
        self.show = show
        self.string = unicode(string)

class CellDetail():
    def __init__(self, colname, colidx=None, rowidx=None):
        self.colname = colname
        self.idx = rowidx
        self.colidx = colidx

class TabWidget(QtGui.QWidget, tab_widget_ui):
    def __init__(self, control, parent=None):
        QtGui.QWidget.__init__(self, parent)
        self.setupUi(self)
        self.control = control

        # HANDLING THE TREEVIEW WIDGET

        self.filterTree.setColumnWidth(0, 234) # COLUMN
        # HANDLING THE CONTEXT MENUS
        self.filterTree.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.columnList.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        self.connect(self.filterTree, QtCore.SIGNAL('customContextMenuRequested(const QPoint&)'),
                     self.on_context_menu_filtertree)
        self.connect(self.columnList, QtCore.SIGNAL('customContextMenuRequested(const QPoint&)'),
                     self.on_context_menu_columnlist)

        self.popCol = QtGui.QMenu()
        self.popFil = QtGui.QMenu()

    def on_context_menu_columnlist(self, point):
        self.popCol.exec_(self.mapToGlobal(point))

    def on_context_menu_filtertree(self, point):
        self.popFil.exec_(self.mapToGlobal(point))

    def closefile(self):
        self.control.filename.close()

    def removecolumnrow(self):

        return

    def removefilterrow(self):
        print "pressed"
        return

    @staticmethod
    def createtreeitem(filterdetail):
        if filterdetail.strict:
            strict = "True"
        else:
            strict = "False"
        if filterdetail.show:
            show = "SHOW"
        else:
            show = "DELETE"
        return QtGui.QTreeWidgetItem([filterdetail.colName, show, filterdetail.string, strict])


class AboutWindow(QtGui.QDialog, dialog_window_ui):
    def __init__(self, parent=None):
        QtGui.QDialog.__init__(self, parent)
        self.setupUi(self)
        self.setGeometry(desktop_width / 3, desktop_height / 3, self.geometry().height(), self.geometry().width())


class MainWindow(QtGui.QMainWindow, main_window_ui):
    def __init__(self, controlitem, parent=None):
        QtGui.QMainWindow.__init__(self, parent)
        self.setupUi(self)
        self.setGeometry(desktop_width / 3, desktop_height / 3, self.geometry().height(), self.geometry().width())

        # MENU BAR HANDLING
        self.actionQuit.activated.connect(self.actionquit_activated)
        self.actionAbout.activated.connect(self.actionabout_activated)
        self.actionFilter.activated.connect(self.addfilter)
        self.actionOpenFile.activated.connect(self.actionopenfile_activated)

        # BUTTONS HANDLING
        self.writeButton.clicked.connect(self.writebutton_clicked)
        self.addFilterButton.clicked.connect(self.addfilter)
        self.removeColumnButton.clicked.connect(self.removecolumnbutton_clicked)

        # TAB WIDGET HANDLING
        # index 0 refers to the first tab
        self.tabList.addTab(TabWidget(controlitem), controlitem.filename.split("/")[-1])
        self.tabList.tabCloseRequested.connect(self.on_tab_close_requested)

    def actionopenfile_activated(self):
        new_file = filenamefromopenfiledialog(self)
        self.tabList.addTab(TabWidget(ControlClass(new_file)), new_file.split("/")[-1])
        return

    def addfilter(self):
        filter_dialog = FilterWindow(self)
        filter_dialog.updatecolcombobox()
        filter_dialog.show()

    def actionquit_activated(self):
        self.close()
        app.quit()

    def on_tab_close_requested(self):
        self.tabList.removeTab(self.tabList.currentIndex())
        # CLOSE APP IF NO TABS ARE OPEN
        if self.tabList.currentIndex() == -1:
            app.quit()

    @staticmethod
    def actionabout_activated():
        about_dialog.show()

    def writebutton_clicked(self):
        file_to_save = filenamefromsavefiledialog(self)
        if len(file_to_save) > 0:
            main_window.tabList.currentWidget().control.writeFile(file_to_save)

    def removecolumnbutton_clicked(self):
        remove = RemoveColumnWindow(self)
        remove.updatecolcombobox()
        remove.show()


class FilterWindow(QtGui.QMainWindow, filter_window_ui):
    def __init__(self, parent=None):
        QtGui.QDialog.__init__(self, parent)
        self.setupUi(self)
        self.control = control
        self.cancelButton.clicked.connect(self.cancelbutton_clicked)
        self.confirmButton.clicked.connect(self.confirmbutton_clicked)
        self.setGeometry(desktop_width / 3, desktop_height / 3, self.geometry().height(), self.geometry().width())

    def cancelbutton_clicked(self):
        self.close()

    def confirmbutton_clicked(self):
        filterdetail = FilterDetails(self.colComboBox.currentText(), self.yesStrictRadio.isChecked(),
                                     self.showRadio.isChecked(), self.filterStringText.toPlainText())
        # CREATING THE TREE
        tree_item = TabWidget.createtreeitem(filterdetail)
        # ADDING IT TO THE TREE
        main_window.tabList.currentWidget().filterTree.addTopLevelItem(tree_item)
        main_window.tabList.currentWidget().control.addfilter(filterdetail)
        self.close()

    def updatecolcombobox(self):
        for colName in main_window.tabList.currentWidget().control.availablecoltitleslist():
            self.colComboBox.addItem(colName)
        return


class RemoveColumnWindow(QtGui.QDialog, remove_column_dialog_ui):
    def __init__(self, parent=None):
        QtGui.QDialog.__init__(self, parent)
        self.setupUi(self)
        self.setGeometry(desktop_width / 3, desktop_height / 3, self.geometry().height(), self.geometry().width())
        self.cancelButton.clicked.connect(self.cancelbutton_clicked)
        self.confirmButton.clicked.connect(self.confirmbutton_clicked)

    def cancelbutton_clicked(self):
        self.close()

    def confirmbutton_clicked(self):
        column = CellDetail(self.colComboBox.currentText())

        # CREATING NEW LIST ITEM
        new_item = QtGui.QListWidgetItem()
        new_item.setText(self.colComboBox.currentText())

        # ADDING IT TO THE LIST
        main_window.tabList.currentWidget().columnList.addItem(new_item)

        main_window.tabList.currentWidget().control.removecolumn(column)
        self.close()

    def updatecolcombobox(self):
        for colName in main_window.tabList.currentWidget().control.availablecoltitleslist():
            self.colComboBox.addItem(colName)
        return


class ControlClass():
    def __init__(self, srcxlsfile):
        self.filename = srcxlsfile
        self.sheet = xlrd.open_workbook(self.filename).sheet_by_index(0)
        self.col_indexes_to_delete = []
        self.row_nums_to_delete = []
        self.col_filter_delete_strict = defaultdict(list)
        self.col_filter_show_strict = defaultdict(list)
        self.col_filter_show_loose = defaultdict(list)
        self.col_filter_delete_loose = defaultdict(list)
        self.polished_file = Workbook()
        self.pf_sheet = self.polished_file.add_sheet(self.sheet.name)

    def cells_with_coltitles(self):
        coltitles = []
        for i in range(0, self.sheet.ncols):
            coltitles.append(CellDetail(self.parseandgetcellvalue(0, i), i, 0))
        return coltitles

    def availablecoltitleslist(self):
        availablecoltitleslist = []
        for cell in self.cells_with_coltitles():
            if cell.colidx not in self.col_indexes_to_delete:
                availablecoltitleslist.append(cell.colname)
        return availablecoltitleslist


    def __colidxfromname__(self, name):
        for cell in self.cells_with_coltitles():
            if cell.colname == name:
                return cell.colidx

    def removecolumn(self, cell):
        self.col_indexes_to_delete.append(self.__colidxfromname__(cell.colname))

    def addfilter(self, filterdetail):
        colidx_to_filter = self.__colidxfromname__(filterdetail.colName)
        if filterdetail.show:
            if filterdetail.strict:
                self.col_filter_show_strict[colidx_to_filter].append(filterdetail.string)
            else:
                self.col_filter_show_loose[colidx_to_filter].append(filterdetail.string)
        else:
            if filterdetail.strict:
                self.col_filter_delete_strict[colidx_to_filter].append(filterdetail.string)
            else:
                self.col_filter_delete_loose[colidx_to_filter].append(filterdetail.string)
        return

    def populaterownumstodelete(self):
        # checking user selected rows
        if len(self.col_filter_delete_strict) > 0:
            for row in range(1, self.sheet.nrows):
                for col in range(0, self.sheet.ncols):
                    if col in self.col_filter_delete_strict and self.parseandgetcellvalue(row, col) in \
                            self.col_filter_delete_strict[col]:
                        self.row_nums_to_delete.append(row)

        # checking strict filter
        if len(self.col_filter_show_strict) > 0:
            for row in range(1, self.sheet.nrows):
                for col in range(0, self.sheet.ncols):
                    if col in self.col_filter_show_strict and not self.parseandgetcellvalue(row, col) in self.col_filter_show_strict[
                        col]:
                        self.row_nums_to_delete.append(row)

        if len(self.col_filter_show_loose) > 0:
            for row in range(1, self.sheet.nrows):
                for col in range(0, self.sheet.ncols):
                    if col in self.col_filter_show_loose:
                        for name in self.col_filter_show_loose[col]:
                            if name not in self.parseandgetcellvalue(row, col):
                                self.row_nums_to_delete.append(row)

        if len(self.col_filter_delete_loose) > 0:
            for row in range(1, self.sheet.nrows):
                for col in range(0, self.sheet.ncols):
                    if col in self.col_filter_delete_loose:
                        for name in self.col_filter_delete_loose[col]:
                            if name in self.parseandgetcellvalue(row, col):
                                self.row_nums_to_delete.append(row)


    def writeFile(self, dstfilename):
        # actual row/col to write since there may be some rows/cols that have to be jumped
        self.populaterownumstodelete()
        col_write = 0
        col_wrote = False
        row_write = 0
        print self.row_nums_to_delete
        for col in range(0, self.sheet.ncols):
            for row in range(0, self.sheet.nrows):
                if col not in self.col_indexes_to_delete and row not in self.row_nums_to_delete:
                    if not col_wrote:
                        col_wrote = True
                    self.pf_sheet.write(row_write, col_write, self.parseandgetcellvalue(row, col))
                    row_write += 1

            if col_write >= (self.sheet.ncols - len(self.col_indexes_to_delete)):
                col_write = 0
            if col_wrote:
                col_write += 1
                col_wrote = False
            row_write = 0

        self.polished_file.save(dstfilename)
        self.polished_file.save(TemporaryFile())

    def parseandgetcellvalue(self, row, col):
        cell = self.sheet.cell(row, col)
        cell_value = cell.value
        # checks whether an value may be a int value and converts it (e.g 2.0 -> 2)
        if cell.ctype in (2,3) and int(cell_value) == cell_value:
            cell_value = int(cell_value)
        return unicode(cell_value)
    
####
# MAIN
####

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    desktop_width = QtGui.QDesktopWidget().geometry().width()
    desktop_height = QtGui.QDesktopWidget().geometry().height()
    filename = filenamefromopenfiledialog()
    if len(filename) > 0:
        control = ControlClass(filename)
        main_window = MainWindow(control)
        about_dialog = AboutWindow(main_window)
        main_window.show()
        app.exec_()
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
import xlrd
import time
from tempfile import TemporaryFile
from collections import defaultdict
from PyQt4 import QtGui, uic, QtCore
from xlwt import Workbook
from xml.etree import cElementTree as ETWrite
from defusedxml import cElementTree as ETRead

main_window_ui = uic.loadUiType("./ui/qt-main.ui")[0]
dialog_window_ui = uic.loadUiType("./ui/qt-about.ui")[0]
filter_window_ui = uic.loadUiType("./ui/qt-filter.ui")[0]
remove_column_dialog_ui = uic.loadUiType("./ui/qt-remove-column.ui")[0]
tab_widget_ui = uic.loadUiType("./ui/qt-tabwidget.ui")[0]


def filename_from_openfile_dialog(parent=None):
    return QtGui.QFileDialog.getOpenFileName(parent, "XLS Polisher - Open Excel File",
                                             "/home",
                                             "Excel Files (*.xls *.xlsx)")


def filename_from_savefile_dialog(parent=None):
    return QtGui.QFileDialog.getSaveFileName(parent, "XLS Polisher - Save Polished File",
                                             "/home",
                                             "Excel Files (*.xls *.xlsx)")


def xml_filename_from_savefile_dialog(parent=None):
    filedialog = QtGui.QFileDialog()
    filedialog.setDefaultSuffix(".xml")
    return filedialog.getSaveFileName(parent, "XLS Polisher - Save Configuration File",
                                      "/home",
                                      "XML Files (*.xml)")


def xml_filename_from_openfile_dialog(parent=None):
    filedialog = QtGui.QFileDialog()
    return filedialog.getOpenFileName(parent, "XLS Polisher - Open Configuration File",
                                      "/home",
                                      "XML Files (*.xml)")


class FilterDetails():
    def __init__(self, colname, strict_bool, show_bool, string):
        self.colName = colname
        self.strict = strict_bool
        self.show = show_bool
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

        self.filterTree.setColumnWidth(0, 234)  # COLUMN
        # HANDLING THE CONTEXT MENUS
        self.popCol = QtGui.QMenu(self.columnList)
        self.popFil = QtGui.QMenu(self.filterTree)
        delete_filter_action = QtGui.QAction("Delete Selected Item", self.popFil)
        delete_column_action = QtGui.QAction("Delete Selected Item", self.popCol)
        self.filterTree.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.columnList.setContextMenuPolicy(QtCore.Qt.ActionsContextMenu)
        self.filterTree.addAction(delete_filter_action)
        self.columnList.addAction(delete_column_action)
        delete_filter_action.triggered.connect(self.on_context_menu_filtertree)
        delete_column_action.triggered.connect(self.on_context_menu_columnlist)


    def on_context_menu_columnlist(self):
        current_item = self.columnList.takeItem(self.columnList.currentRow())
        if current_item:
            self.control.remove_item_from_list_item(current_item)

    def on_context_menu_filtertree(self):
        current_item = self.filterTree.takeTopLevelItem(self.filterTree.indexOfTopLevelItem(self.filterTree.currentItem()))
        if current_item:
            self.control.remove_item_from_tree_item(current_item)

    def closefile(self):
        self.control.filename.close()

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
        self.actionQuit.activated.connect(self.on_actionquit_activated)
        self.actionAbout.activated.connect(self.on_actionabout_activated)
        self.actionFilter.activated.connect(self.addfilter)
        self.actionOpenFile.activated.connect(self.on_actionopenfile_activated)
        self.actionOpen_Configuration.activated.connect(self.on_open_configuration_activated)
        self.actionSave_Configuration.activated.connect(self.on_save_configuration_activated)

        # BUTTONS HANDLING
        self.writeButton.clicked.connect(self.writebutton_clicked)
        self.addFilterButton.clicked.connect(self.addfilter)
        self.removeColumnButton.clicked.connect(self.removecolumnbutton_clicked)

        # TAB WIDGET HANDLING
        # index 0 refers to the first tab
        self.tabList.addTab(TabWidget(controlitem), controlitem.filename.split("/")[-1])
        self.tabList.tabCloseRequested.connect(self.on_tab_close_requested)

    def on_actionopenfile_activated(self):
        new_file = filename_from_openfile_dialog(self)
        if len(new_file) > 0:
            self.tabList.addTab(TabWidget(ControlClass(new_file)), new_file.split("/")[-1])
        return

    def addfilter(self):
        filter_dialog = FilterWindow(self)
        filter_dialog.updatecolcombobox()
        filter_dialog.show()

    def on_actionquit_activated(self):
        self.close()
        app.quit()

    def on_tab_close_requested(self):
        self.tabList.removeTab(self.tabList.currentIndex())
        # CLOSE APP IF NO TABS ARE OPEN
        if self.tabList.currentIndex() == -1:
            app.quit()

    @staticmethod
    def on_actionabout_activated():
        about_dialog.show()

    def writebutton_clicked(self):
        file_to_save = filename_from_savefile_dialog(self)
        if len(file_to_save) > 0:
            self.tabList.currentWidget().control.writeFile(file_to_save)
            self.tabList.currentWidget().control.renewWorkbook()

    def removecolumnbutton_clicked(self):
        remove = RemoveColumnWindow(self)
        remove.updatecolcombobox()
        remove.show()

    def on_save_configuration_activated(self):
        self.tabList.currentWidget().control.createConfFile(self.tabList.currentWidget())
        return

    def on_open_configuration_activated(self):
        filename = xml_filename_from_openfile_dialog()
        if len(filename) > 0:
            self.tabList.currentWidget().control.loadConfFile(filename, self.tabList.currentWidget())
        return


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
        self.workbook, self.wb_sheet = self.new_workbook_and_workbook_sheet(self.sheet.name)

    def new_workbook_and_workbook_sheet(self, sheet_name):
        workbook = Workbook()
        workbook_sheet = workbook.add_sheet(sheet_name)
        return workbook, workbook_sheet

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

    def must_delete(self, value, col):
        # CHECKING STRICT FILTERS

        if self.col_filter_delete_strict[col]:
            if value in self.col_filter_delete_strict[col]:
                return True

        if self.col_filter_show_strict[col]:
            if value not in self.col_filter_show_strict[col]:
                return True

        # CHECKING LOOSE FILTERS
        # split() lenght will be 1 if the value isn't contained in our string

        if self.col_filter_show_loose[col]:
            for string in self.col_filter_show_loose[col]:
                if len(value.lower().split(string.lower())) > 1:
                    return False
            return True

        if self.col_filter_delete_loose[col]:
            for string in self.col_filter_delete_loose[col]:
                if len(value.lower().split(string.lower())) > 1:
                    return False
            return True
        return False

    def populaterownumstodelete(self):
        cols_to_check = (set(self.col_filter_delete_strict)
                         .union(self.col_filter_show_strict)
                         .union(self.col_filter_show_loose)
                         .union(self.col_filter_delete_loose)
                         .intersection(range(self.sheet.ncols)))
        for row in range(1, self.sheet.nrows):
            for col in cols_to_check:
                value = self.parseandgetcellvalue(row, col)
                if self.must_delete(value, col):
                    self.row_nums_to_delete.append(row)
                    break

    def writeFile(self, dstfilename):
        # actual row/col to write since there may be some rows/cols that have to be jumped
        self.populaterownumstodelete()
        col_write = 0
        col_wrote = False
        row_write = 0
        for col in (cols for cols in range(self.sheet.ncols) if cols not in self.col_indexes_to_delete):
            for row in (rows for rows in range(self.sheet.nrows) if rows not in self.row_nums_to_delete):
                if not col_wrote:
                    col_wrote = True
                self.wb_sheet.write(row_write, col_write, self.parseandgetcellvalue(row, col))
                row_write += 1

            if col_write >= (self.sheet.ncols - len(self.col_indexes_to_delete)):
                col_write = 0
            if col_wrote:
                col_write += 1
                col_wrote = False
            row_write = 0

        # CLEARING LIST
        del (self.row_nums_to_delete[:])
        self.workbook.save(dstfilename)
        self.workbook.save(TemporaryFile())

    def parseandgetcellvalue(self, row, col):
        cell = self.sheet.cell(row, col)
        cell_value = cell.value
        # checks whether an value may be a int value and converts it (e.g 2.0 -> 2)
        if cell.ctype in (2, 3) and int(cell_value) == cell_value:
            cell_value = int(cell_value)
        return unicode(cell_value)

    def createConfFile(self, tabListWidget):
        root_name = "data"
        xml_root = ETWrite.Element(root_name)
        filter_name = "filter"
        # ADDING FILTER ELEMENTS
        # GETS THE FIRST ITEM
        xml_filter = ETWrite.SubElement(xml_root, filter_name)
        filteritem = tabListWidget.filterTree.topLevelItem(0)
        while filteritem is not None:
            xml_filter_item = ETWrite.SubElement(xml_filter, "filter_item")

            xml_filter_item_detail_column = ETWrite.SubElement(xml_filter_item, "column")
            xml_filter_item_detail_column.text = unicode(filteritem.text(0))

            xml_filter_item_detail_mode = ETWrite.SubElement(xml_filter_item, "mode")
            xml_filter_item_detail_mode.text = unicode(filteritem.text(1))

            xml_filter_item_detail_filter = ETWrite.SubElement(xml_filter_item, "filter")
            xml_filter_item_detail_filter.text = unicode(filteritem.text(2))

            xml_filter_item_detail_strict = ETWrite.SubElement(xml_filter_item, "strict")
            xml_filter_item_detail_strict.text = unicode(filteritem.text(3))

            filteritem = tabListWidget.filterTree.itemBelow(filteritem)

        # ADDING COLUMN REMOVAL ELEMENTS
        delete_column_name = "columndelete"
        xml_deletecolumn = ETWrite.SubElement(xml_root, delete_column_name)
        for index in range(tabListWidget.columnList.count()):
            delete_column_item = tabListWidget.columnList.item(index)
            ETWrite.SubElement(xml_deletecolumn, "column", {"name": unicode(delete_column_item.text())})

        ETWrite.ElementTree(xml_root).write(xml_filename_from_savefile_dialog())

    def loadConfFile(self, filename, tabListWidget):
        tree = ETRead.parse(filename)
        filterTree = tabListWidget.filterTree
        columnList = tabListWidget.columnList
        root = tree.getroot()

        # POPULATING FILTERS
        for filter in (filters for filters in root if filters.tag == "filter"):
            for item in filter.findall('filter_item'):
                column = item.find("column").text
                show = item.find("mode").text
                if show.lower() == "show":
                    show = True
                else:
                    show = False
                filterstring = item.find("filter").text
                # HOTFIX FOR EMPTY STRINGS
                if filterstring is None:
                    filterstring = ""
                strict = item.find("strict").text
                if strict.lower() == "true":
                    strict = True
                else:
                    strict = False

                filter_detail = FilterDetails(column, strict, show, filterstring)
                tree_item = TabWidget.createtreeitem(filter_detail)
                filterTree.addTopLevelItem(tree_item)
                tabListWidget.control.addfilter(filter_detail)


        # POPULATING COLUMNS TO DELETE
        for column in (columns for columns in root if columns.tag == "columndelete"):
            for column_item in column:
                column = column_item.get('name')
                # CREATING NEW LIST ITEM
                new_item = QtGui.QListWidgetItem()
                new_item.setText(column)
                # ADDING IT TO THE LIST
                columnList.addItem(new_item)
                main_window.tabList.currentWidget().control.removecolumn(CellDetail(column))

        return

    def remove_item_from_list_item(self, item):
        if item:
            column = item.text()
            self.col_indexes_to_delete.remove(self.__colidxfromname__(column))

    def remove_filterdetail_from_list(self, filterdetail):
        colidx = self.__colidxfromname__(filterdetail.colName)
        if filterdetail.strict and filterdetail.show:
            def_dict = self.col_filter_show_strict
        elif filterdetail.strict and not filterdetail.show:
            def_dict = self.col_filter_delete_strict
        elif not filterdetail.strict and filterdetail.show:
            def_dict = self.col_filter_show_loose
        elif not filterdetail.strict and not filterdetail.show:
            def_dict = self.col_filter_delete_loose
        def_dict[colidx].remove(filterdetail.string)
        if not def_dict[colidx]:
            del(def_dict[colidx])
        return

    def remove_item_from_tree_item(self, item):
        if item:
            column = unicode(item.text(0))
            mode = unicode(item.text(1))
            filterstring = unicode(item.text(2))
            strict = unicode(item.text(3))
            filter_detail = self.filterdetail_from_strings(column, strict, mode, filterstring)
            self.remove_filterdetail_from_list(filter_detail)

    def filterdetail_from_strings(self, colname, strict, mode, string):
        if strict.lower() in ["true"]:
            strict_bool = True
        else:
            strict_bool = False
        if mode.lower() in ["show"]:
            show_bool = True
        else:
            show_bool = True
        return FilterDetails(colname, strict_bool, show_bool, string)

    def renewWorkbook(self):
        self.workbook, self.wb_sheet = self.new_workbook_and_workbook_sheet(self.sheet.name)

####
# MAIN
####

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    desktop_width = QtGui.QDesktopWidget().geometry().width()
    desktop_height = QtGui.QDesktopWidget().geometry().height()
    filename = filename_from_openfile_dialog()
    if len(filename) > 0:
        control = ControlClass(filename)
        main_window = MainWindow(control)
        about_dialog = AboutWindow(main_window)
        main_window.show()
        app.exec_()
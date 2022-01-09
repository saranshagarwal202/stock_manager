from MainWindow import Ui_MainWindow

from PyQt5 import QtWidgets as qtw
from PyQt5 import QtCore as qtc
from PyQt5 import QtGui as qtg

from copy import copy, deepcopy
import pandas as pd
import json
import os


class MainWindow(qtw.QMainWindow):

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        self.ui.actionSave_As.setEnabled(False)
        self.ui.actionSave.setEnabled(False)
        self.ui.tabWidget.setTabEnabled(3, False)
        self.ui.tabWidget.setTabEnabled(2, False)
        self.ui.tabWidget.setTabEnabled(1, False)
        self.ui.tabWidget.setTabEnabled(0, False)

        # Push buttons
        ##### Stock Sheet #####
        # self.ui.edit_button.setEnabled(False)
        self.ui.clear_button.setEnabled(False)
        self.ui.delete_entry.setEnabled(False)
        self.ui.reset_find_preview.setEnabled(False)
        self.ui.confirm_edit.setEnabled(False)
        self.ui.confirm_edit.clicked.connect(self.confirm_edit)
        self.ui.Add_stock.clicked.connect(self.add_stock)
        self.ui.find_button.clicked.connect(self.find_item_in_stock_sheet)
        self.ui.delete_entry.clicked.connect(self.delete_entry_from_tabel)
        self.ui.clear_button.clicked.connect(
            lambda: self.clear_tabel(self.ui.current_stock_find_entry))
        self.ui.clear_button.clicked.connect(
            lambda: self.clear_tabel(self.ui.find_stock_preview, True))
        self.ui.clear_button.clicked.connect(self.stock_tab_button_set_False)
        self.ui.reset_find_preview.clicked.connect(self.reset_preview)
        self.ui.current_stock_find_entry.keyPressEvent = self.find_enter_event
        self.ui.find_stock_preview.keyPressEvent = self.preview_enter_event
        self.ui.scrollArea_2.mousePressEvent = self.reset_selected

        ##### Issue Sheet #####
        self.ui.clear_button_2.setEnabled(False)
        self.ui.delete_entry_2.setEnabled(False)
        self.ui.edit_button_2.setEnabled(False)
        self.ui.Reset_issued_preview.setEnabled(False)
        self.ui.approval_issue.setEnabled(False)
        self.ui.approval_sale.setEnabled(False)
        self.ui.add_sku.setEnabled(False)
        self.ui.find_button_2.clicked.connect(
            lambda: self.find_item_in_stock_sheet(sheet="issue_sheet"))
        self.ui.clear_button_2.clicked.connect(
            lambda: self.clear_tabel(self.ui.issued_stock_find_entry))
        self.ui.clear_button_2.clicked.connect(
            lambda: self.clear_tabel(self.ui.issue_stock_preview, True))
        self.ui.clear_button_2.clicked.connect(self.issue_tab_button_set_false)
        # tabel
        # self.ui.find_stock_preview.clicked.connect(lambda : self.ui.delete_entry.setEnabled(True))

        # Variabels
        self.saved = True
        self.deleted = []
        self.original = []
        self.original_prev = []

        # Action buttons
        self.ui.actionOpen_Project.triggered.connect(self.open_project)
        self.ui.actionExit_2.triggered.connect(self.exit)
        self.ui.actionNew_Project.triggered.connect(self.New_project)
        self.ui.actionSave.triggered.connect(
            lambda: self.save_project(save_path=self.project_path))
        self.ui.actionSave_As.triggered.connect(lambda: self.save_project(qtw.QFileDialog.getExistingDirectory(
            self, "Select Directory")))

    def add_stock(self):
        entry_dict = {}
        for j in range(self.ui.current_entry_add_stock.columnCount()):
            try:
                entry_dict[self.ui.current_entry_add_stock.horizontalHeaderItem(
                    j).text()] = self.ui.current_entry_add_stock.item(0, j).text()
            except AttributeError:
                entry_dict[self.ui.current_entry_add_stock.horizontalHeaderItem(
                    j).text()] = self.ui.current_entry_add_stock.item(0, j)
        self.stock_sheet.append(entry_dict)
        # self.ui.current_entry_add_stock.clear()
        self.load_data_in_tabel(self.ui.current_stock, entry_dict, len(
            self.stock_sheet), list(entry_dict), len(self.stock_sheet)-1)
        tabel_columns = []
        for i in range(self.ui.current_entry_add_stock.columnCount()):
            tabel_columns.append(
                self.ui.current_entry_add_stock.horizontalHeaderItem(i).text())
        self.ui.current_entry_add_stock.clear()
        self.ui.current_entry_add_stock.setHorizontalHeaderLabels(
            tabel_columns)

    def confirm_edit(self):
        for i in range(self.ui.find_stock_preview.rowCount()):
            for j in range(1, self.ui.find_stock_preview.columnCount()):

                item = self.ui.find_stock_preview.item(i, j)

                self.stock_sheet[int(self.ui.find_stock_preview.item(i, 0).text(
                ))-1][self.ui.find_stock_preview.horizontalHeaderItem(j).text()] = item.text()
                self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                    self.stock_sheet), self.schema["stock_sheet"]["columns"])

    def reset_selected(self, event):
        self.ui.current_stock_find_entry.clearSelection()
        self.ui.find_stock_preview.clearSelection()
        self.ui.current_entry_add_stock.clearSelection()
        self.ui.current_stock.clearSelection()

    def issue_find_enter_event(self, event):
        # Enter key pressed
        if event.key() == qtc.Qt.Key_Return or event.key() == qtc.Qt.Key_Enter:
            if len(self.ui.issued_stock_find_entry.selectedIndexes()) == 1:
                item = self.ui.issued_stock_find_entry.itemFromIndex(
                    self.ui.issued_stock_find_entry.selectedIndexes()[0])
                if item is None:
                    self.ui.issued_stock_find_entry.edit(
                        self.ui.issued_stock_find_entry.selectedIndexes()[0])
                else:
                    self.find_item_in_stock_sheet()
        qtw.QTableWidget.keyPressEvent(self.ui.issued_stock_find_entry, event)
    
    def issue_preview_enter_event(self, event):
        # Enter key pressed
        if event.key() == qtc.Qt.Key_Return or event.key() == qtc.Qt.Key_Enter:
            index = self.ui.issue_stock_preview.selectedIndexes()
            if len(index) == 1:
                index = index[0]
                item = self.ui.issue_stock_preview.itemFromIndex(index)

                self.stock_sheet[int(self.ui.issue_stock_preview.item(index.row(
                ), 0).text())-1][self.ui.issue_stock_preview.horizontalHeaderItem(index.column()).text()] = item.text()
                self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                    self.stock_sheet), self.schema["stock_sheet"]["columns"])
        if event.key() == qtc.Qt.Key_Delete:
            self.delete_entry_from_tabel()

        qtw.QTableWidget.keyPressEvent(self.ui.issue_stock_preview, event)

    def find_enter_event(self, event):
        # Enter key pressed
        if event.key() == qtc.Qt.Key_Return or event.key() == qtc.Qt.Key_Enter:
            if len(self.ui.current_stock_find_entry.selectedIndexes()) == 1:
                item = self.ui.current_stock_find_entry.itemFromIndex(
                    self.ui.current_stock_find_entry.selectedIndexes()[0])
                if item is None:
                    self.ui.current_stock_find_entry.edit(
                        self.ui.current_stock_find_entry.selectedIndexes()[0])
                else:
                    self.find_item_in_stock_sheet()
        qtw.QTableWidget.keyPressEvent(self.ui.current_stock_find_entry, event)

    def preview_enter_event(self, event):
        # Enter key pressed
        if event.key() == qtc.Qt.Key_Return or event.key() == qtc.Qt.Key_Enter:
            index = self.ui.find_stock_preview.selectedIndexes()
            if len(index) == 1:
                index = index[0]
                item = self.ui.find_stock_preview.itemFromIndex(index)

                self.stock_sheet[int(self.ui.find_stock_preview.item(index.row(
                ), 0).text())-1][self.ui.find_stock_preview.horizontalHeaderItem(index.column()).text()] = item.text()
                self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                    self.stock_sheet), self.schema["stock_sheet"]["columns"])
        if event.key() == qtc.Qt.Key_Delete:
            self.delete_entry_from_tabel()

        qtw.QTableWidget.keyPressEvent(self.ui.find_stock_preview, event)

    def reset_preview(self):
        for dicti in self.original:

            index = copy(int(dicti["sheet_index"]))
            temp_dict = deepcopy(dicti)
            del(temp_dict["sheet_index"])
            self.stock_sheet[index-1] = deepcopy(temp_dict)
        self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
            self.stock_sheet), self.schema["stock_sheet"]["columns"])
        self.load_data_in_tabel(self.ui.find_stock_preview, self.original, len(
            self.original), list(self.original[0]))

    def clear_tabel(self, tabel, disable=False):

        assert(isinstance(tabel, qtw.QTableWidget))
        tabel_columns = []
        for i in range(tabel.columnCount()):
            tabel_columns.append(tabel.horizontalHeaderItem(i).text())
        tabel.clear()
        tabel.setHorizontalHeaderLabels(tabel_columns)
        tabel.setRowCount(1)
        if disable == True:
            tabel.setEnabled(False)

    def stock_tab_button_set_False(self):
        self.ui.clear_button.setEnabled(False)
        self.ui.delete_entry.setEnabled(False)
        self.ui.reset_find_preview.setEnabled(False)
        self.ui.confirm_edit.setEnabled(False)

    def issue_tab_button_set_false(self):
        self.ui.clear_button_2.setEnabled(False)
        self.ui.delete_entry_2.setEnabled(False)
        self.ui.edit_button_2.setEnabled(False)
        self.ui.Reset_issued_preview.setEnabled(False)

    def save_edit(self):
        for i in range(self.ui.find_stock_preview.rowCount()):
            temp_dict = {}
            for j in range(self.ui.find_stock_preview.columnCount()):
                temp_dict[self.ui.find_stock_preview.horizontalHeaderItem(
                    j).text()] = self.ui.find_stock_preview.item(i, j).text()
            index = copy(int(temp_dict["sheet_index"]))
            del(temp_dict["sheet_index"])
            self.stock_sheet[index-1] = deepcopy(temp_dict)
        self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
            self.stock_sheet), self.schema["stock_sheet"]["columns"])

    def delete_entry_from_tabel(self):
        if len(self.ui.find_stock_preview.selectedIndexes()) == self.ui.find_stock_preview.columnCount():
            # save item
            temp_dict = {}
            for i in range(self.ui.find_stock_preview.columnCount()):
                temp_dict[self.ui.find_stock_preview.horizontalHeaderItem(i).text()] = self.ui.find_stock_preview.item(
                    self.ui.find_stock_preview.selectedIndexes()[0].row(), i).text()
            self.deleted.append(self.stock_sheet.pop(
                int(temp_dict["sheet_index"])-1))
            self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                self.stock_sheet), self.schema["stock_sheet"]["columns"])
            self.ui.find_stock_preview.removeRow(
                self.ui.find_stock_preview.selectedIndexes()[0].row())
        else:
            indexes = self.ui.find_stock_preview.selectedIndexes()
            for index in indexes:
                self.ui.find_stock_preview.setItem(
                    index.row(), index.column(), qtw.QTableWidgetItem(str('')))
                self.stock_sheet[int(self.ui.find_stock_preview.item(index.row(
                ), 0).text())-1][self.ui.find_stock_preview.horizontalHeaderItem(index.column()).text()] = ''
            self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
                self.stock_sheet), self.schema["stock_sheet"]["columns"])

            # self.deleted.append(self.ui.)
        # print(self.deleted)

    def find_item_in_stock_sheet(self, sheet="stock_sheet"):

        if sheet == "stock_sheet":
            self.clear_tabel(self.ui.find_stock_preview)
            self.original = []
            self.ui.current_stock_find_entry.clearSelection()
            # sku_code = self.ui.sku_code.text()
            # if ''.join(sku_code.split()) == "":
            search_dict = {}
            for i in range(len(self.schema["stock_sheet"]["columns"])):
                try:
                    # print(self.ui.current_stock_find_entry.item(0,i).text())
                    if self.ui.current_stock_find_entry.item(0, i).text() != "":
                        search_dict[self.schema["stock_sheet"]["columns"][i]
                                    ] = self.ui.current_stock_find_entry.item(0, i).text()
                except AttributeError:
                    False
                #     print("in exception")
            if not bool(search_dict):
                return
            # search_result = []
            count = 0
            # print(self.ui.current_stock.items())

            for dicti in self.stock_sheet:
                count += 1
                found_flag = True
                for key in list(search_dict):
                    if search_dict[key] != str(dicti[key]):
                        found_flag = False
                        break
                if found_flag == True:
                    dicti_temp = {"sheet_index": count}
                    dicti_temp.update(dicti)
                    self.original.append(dicti_temp)
            if bool(self.original):
                self.ui.find_stock_preview.setEnabled(True)
                self.load_data_in_tabel(self.ui.find_stock_preview, self.original, len(
                    self.original), list(self.original[0]))
                self.ui.clear_button.setEnabled(True)
                self.ui.reset_find_preview.setEnabled(True)
                self.ui.confirm_edit.setEnabled(True)
                # self.ui.edit_button.setEnabled(True)
                self.ui.delete_entry.setEnabled(True)

        elif sheet == "issue_sheet":
            self.clear_tabel(self.ui.issue_stock_preview)
            self.original = []
            self.ui.issued_stock_find_entry.clearSelection()
            # sku_code = self.ui.sku_code.text()
            # if ''.join(sku_code.split()) == "":
            search_dict = {}
            for i in range(len(self.schema["issue_sheet"]["columns"])):
                try:
                    # print(self.ui.current_stock_find_entry.item(0,i).text())
                    if self.ui.issued_stock_find_entry.item(0, i).text() != "":
                        search_dict[self.schema["issue_sheet"]["columns"][i]
                                    ] = self.ui.issued_stock_find_entry.item(0, i).text()
                except AttributeError:
                    False
                #     print("in exception")
            if not bool(search_dict):
                return
            # search_result = []
            count = 0
            # print(self.ui.current_stock.items())

            for dicti in self.issue_sheet:
                count += 1
                found_flag = True
                for key in list(search_dict):
                    if search_dict[key] != str(dicti[key]):
                        found_flag = False
                        break
                if found_flag == True:
                    dicti_temp = {"sheet_index": count}
                    dicti_temp.update(dicti)
                    self.original.append(dicti_temp)
            if bool(self.original):
                self.ui.issue_stock_preview.setEnabled(True)
                self.load_data_in_tabel(self.ui.issue_stock_preview, self.original, len(
                    self.original), list(self.original[0]))
                self.ui.clear_button_2.setEnabled(True)
                self.ui.Reset_issued_preview.setEnabled(True)
                self.ui.edit_button_2.setEnabled(True)
                self.ui.delete_entry_2.setEnabled(True)

    def load_data_in_tabel(self, tabel, values, rowcount, columns, affected_row=None):
        # affected_row =
        assert(isinstance(tabel, qtw.QTableWidget))
        tabel.setRowCount(rowcount)
        tabel.setColumnCount(
            len(columns))
        tabel.setHorizontalHeaderLabels(
            columns)
        if affected_row == None:
            # load entire data in table
            for row, dicti in enumerate(values):
                for col, key in enumerate(dicti):
                    tabel.setItem(
                        row, col, qtw.QTableWidgetItem(str(dicti[key])))
        else:
            # load only the respective data

            for col, key in enumerate(values):
                # here values is one dict
                tabel.setItem(affected_row, col,
                              qtw.QTableWidgetItem(values[key]))

    def save_project(self, save_path):
        self.stock_sheet.to_excel(os.path.join(
            save_path, self.schema["stock_sheet"]["filename"]))
        self.issue_sheet.to_excel(os.path.join(
            save_path, self.schema["issue_sheet"]["filename"]))
        self.issue_return_sheet.to_excel(os.path.join(
            save_path, self.schema["issue_return_sheet"]["filename"]))
        self.sale_sheet.to_excel(os.path.join(
            save_path, self.schema["sale_sheet"]["filename"]))

    def open_project(self):
        # try:
        self.project_path = qtw.QFileDialog.getExistingDirectory(
            self, "Select Directory")
        f = open(os.path.join(self.project_path, "schema.json"))
        self.schema = json.loads(f.read())
        f.close()

        self.stock_sheet = pd.read_excel(os.path.join(
            self.project_path, self.schema["stock_sheet"]["filename"])).to_dict(orient="records")
        self.issue_sheet = pd.read_excel(os.path.join(
            self.project_path, self.schema["issue_sheet"]["filename"])).to_dict(orient="records")
        self.issue_return_sheet = pd.read_excel(os.path.join(
            self.project_path, self.schema["issue_return_sheet"]["filename"])).to_dict(orient="records")
        self.sale_sheet = pd.read_excel(os.path.join(
            self.project_path, self.schema["sale_sheet"]["filename"])).to_dict(orient="records")

        # stock tab
        self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
            self.stock_sheet), self.schema["stock_sheet"]["columns"])
        self.load_data_in_tabel(self.ui.current_stock_find_entry, [
        ], 1, self.schema["stock_sheet"]["columns"])
        self.load_data_in_tabel(
            self.ui.find_stock_preview, [], 1, self.schema["stock_sheet"]["columns"])
        self.load_data_in_tabel(self.ui.current_entry_add_stock, [
        ], 1, self.schema["stock_sheet"]["columns"])

        # issue tab
        self.load_data_in_tabel(self.ui.issued_stock_find_entry, [
        ], 1, self.schema["issue_sheet"]["columns"])
        self.load_data_in_tabel(
            self.ui.issue_stock_preview, [], 1, self.schema["issue_sheet"]["columns"])
        self.load_data_in_tabel(self.ui.sku_find, [], 1,
                                self.schema["stock_sheet"]["columns"])
        self.load_data_in_tabel(self.ui.issue_items, [],
                                1, self.schema["issue_sheet"]["columns"])

        # qtw.QHeaderView.setSectionResizeMode(qtw.QHeaderView.Stretch)

        self.ui.tabWidget.setTabEnabled(3, True)
        self.ui.tabWidget.setTabEnabled(2, True)
        self.ui.tabWidget.setTabEnabled(1, True)
        self.ui.tabWidget.setTabEnabled(0, True)
        self.ui.actionSave_As.setEnabled(True)
        self.ui.actionSave.setEnabled(True)

        self.ui.issued_stock_find_entry.horizontalHeader(
        ).resizeSections(qtw.QHeaderView.ResizeToContents)
        self.ui.issue_stock_preview.horizontalHeader(
        ).resizeSections(qtw.QHeaderView.ResizeToContents)
        self.ui.sku_find.horizontalHeader().resizeSections(
            qtw.QHeaderView.ResizeToContents)
        self.ui.issue_items.horizontalHeader().resizeSections(
            qtw.QHeaderView.ResizeToContents)

        self.ui.current_stock_find_entry.horizontalHeader(
        ).resizeSections(qtw.QHeaderView.ResizeToContents)
        self.ui.find_stock_preview.horizontalHeader().resizeSections(
            qtw.QHeaderView.ResizeToContents)
        self.ui.current_entry_add_stock.horizontalHeader(
        ).resizeSections(qtw.QHeaderView.ResizeToContents)
        self.ui.current_stock.horizontalHeader().resizeSections(
            qtw.QHeaderView.ResizeToContents)

        # except:
        #     msg = qtw.QMessageBox()
        #     msg.setText("Unable to open project, please check path/schema")
        #     msg.setStandardButtons(qtw.QMessageBox.Ok)
        #     msg.setIcon(qtw.QMessageBox.Warning)
        #     msg.setModal(True)
        #     msg.exec()

    def New_project(self):
        self.project_path = qtw.QFileDialog.getExistingDirectory(
            self, "Select Directory")
        f = open("schema.json")
        self.schema = json.loads(f.read())
        f.close()
        f = open(os.path.join(self.project_path, "schema.json"), "w")
        f.write(json.dumps(self.schema))
        pd.DataFrame(columns=self.schema["stock_sheet"]["columns"]).to_excel(os.path.join(
            self.project_path, self.schema["stock_sheet"]["filename"]), index=False)
        pd.DataFrame(columns=self.schema["issue_sheet"]["columns"]).to_excel(os.path.join(
            self.project_path, self.schema["issue_sheet"]["filename"]), index=False)
        pd.DataFrame(columns=self.schema["issue_return_sheet"]["columns"]).to_excel(os.path.join(
            self.project_path, self.schema["issue_return_sheet"]["filename"]), index=False)
        pd.DataFrame(columns=self.schema["sale_sheet"]["columns"]).to_excel(os.path.join(
            self.project_path, self.schema["sale_sheet"]["filename"]), index=False)

        self.ui.tabWidget.setTabEnabled(3, True)
        self.ui.tabWidget.setTabEnabled(2, True)
        self.ui.tabWidget.setTabEnabled(1, True)
        self.ui.tabWidget.setTabEnabled(0, True)
        self.ui.actionSave_As.setEnabled(True)
        self.ui.actionSave.setEnabled(True)

    def exit(self):
        if self.saved == False:
            msg = qtw.QMessageBox()
            msg.setText("Please save before closing")
            msg.setStandardButtons(qtw.QMessageBox.Save |
                                   qtw.QMessageBox.Discard)
            msg.setModal(True)
            msg.buttonClicked.connect(self.save_warning)
            msg.exec_()

            # print(retval)
        else:
            os._exit(0)

    def save_warning(self, item):
        if item.text() == "Save":
            self.save_project(self.project_path)


if __name__ == "__main__":

    app = qtw.QApplication([])

    window = MainWindow()
    window.showMaximized()

    app.exec()

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
        self.ui.edit_button.setEnabled(False)
        self.ui.clear_button.setEnabled(False)
        self.ui.delete_entry.setEnabled(False)
        self.ui.reset_find_preview.setEnabled(False)
        # self.ui.sku_code.returnPressed.connect(self.find_item_in_stock_sheet)
        self.ui.find_button.clicked.connect(self.find_item_in_stock_sheet)
        self.ui.delete_entry.clicked.connect(self.delete_entry_from_tabel)
        self.ui.edit_button.clicked.connect(self.save_edit)
        self.ui.clear_button.clicked.connect(lambda: self.clear_tabel(self.ui.current_stock_find_entry))
        self.ui.clear_button.clicked.connect(lambda: self.clear_tabel(self.ui.find_stock_preview, True))
        self.ui.clear_button.clicked.connect(lambda: self.ui.clear_button.setEnabled(False))
        self.ui.reset_find_preview.clicked.connect(self.reset_preview)

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

    def reset_preview(self):
        for dicti in self.original:
            
            index = copy(int(dicti["sheet_index"]))
            temp_dict = deepcopy(dicti)
            del(temp_dict["sheet_index"])
            self.stock_sheet[index-1] = deepcopy(temp_dict)
        self.load_data_in_tabel(self.ui.current_stock,self.stock_sheet, len(self.stock_sheet), self.schema["stock_sheet"]["columns"])
        self.load_data_in_tabel(self.ui.find_stock_preview,self.original, len(self.original), list(self.original[0]))

    def clear_tabel(self, tabel, disable = False):
        
        assert(isinstance(tabel, qtw.QTableWidget))
        tabel_columns = []
        for i in  range(tabel.columnCount()):
            tabel_columns.append(tabel.horizontalHeaderItem(i).text())
        tabel.clear()
        tabel.setHorizontalHeaderLabels(tabel_columns)
        tabel.setRowCount(1)
        if disable == True:
            tabel.setEnabled(False)



    def save_edit(self):
        for i in range(self.ui.find_stock_preview.rowCount()):
            temp_dict = {}
            for j in range(self.ui.find_stock_preview.columnCount()):
                temp_dict[self.ui.find_stock_preview.horizontalHeaderItem(j).text()] = self.ui.find_stock_preview.item(i,j).text()
            index = copy(int(temp_dict["sheet_index"]))
            del(temp_dict["sheet_index"])
            self.stock_sheet[index-1] = deepcopy(temp_dict)
        self.load_data_in_tabel(self.ui.current_stock,self.stock_sheet, len(self.stock_sheet), self.schema["stock_sheet"]["columns"])

    def delete_entry_from_tabel(self):
        if len(self.ui.find_stock_preview.selectedIndexes())==self.ui.find_stock_preview.columnCount():
            # save item 
            temp_dict = {}
            for i in range(self.ui.find_stock_preview.columnCount()):
                temp_dict[self.ui.find_stock_preview.horizontalHeaderItem(i).text()] = self.ui.find_stock_preview.item(self.ui.find_stock_preview.selectedIndexes()[0].row(),i).text()
            self.deleted.append(self.stock_sheet.pop(int(temp_dict["sheet_index"])-1))
            self.load_data_in_tabel(self.ui.current_stock,self.stock_sheet, len(self.stock_sheet), self.schema["stock_sheet"]["columns"])
            self.ui.find_stock_preview.removeRow(self.ui.find_stock_preview.selectedIndexes()[0].row())

            # self.deleted.append(self.ui.)
        print(self.deleted)


    def find_item_in_stock_sheet(self):
        
        self.original = []
        self.ui.current_stock_find_entry.clearSelection()
        # sku_code = self.ui.sku_code.text()
        # if ''.join(sku_code.split()) == "":
        search_dict = {}
        for i in range(len(self.schema["stock_sheet"]["columns"])):
            try:
            # print(self.ui.current_stock_find_entry.item(0,i).text())
                if self.ui.current_stock_find_entry.item(0,i).text() != "":
                    search_dict[self.schema["stock_sheet"]["columns"][i]
                                ] = self.ui.current_stock_find_entry.item(0,i).text()
            except AttributeError:
                False
            #     print("in exception")
        if not bool(search_dict):
            return
        # search_result = []
        count = 0
        # print(self.ui.current_stock.items())

        for dicti in self.stock_sheet:
            count+=1
            found_flag = True
            for key in list(search_dict):
                if search_dict[key] != dicti[key]:
                    found_flag = False
                    break
            if found_flag == True:
                dicti_temp = {"sheet_index": count}
                dicti_temp.update(dicti)
                self.original.append(dicti_temp)
        if bool(self.original):
            self.ui.find_stock_preview.setEnabled(True)
            self.load_data_in_tabel(self.ui.find_stock_preview, self.original, len(self.original), list(self.original[0]))
            self.ui.clear_button.setEnabled(True)
            self.ui.reset_find_preview.setEnabled(True)
            self.ui.edit_button.setEnabled(True)
            self.ui.delete_entry.setEnabled(True)
            
                

    def load_data_in_tabel(self, tabel, values, rowcount, columns):
        assert(isinstance(tabel, qtw.QTableWidget))
        tabel.setRowCount(rowcount)
        tabel.setColumnCount(
            len(columns))
        tabel.setHorizontalHeaderLabels(
            columns)

        for row, dicti in enumerate(values):
            for col, key in enumerate(dicti):
                tabel.setItem(
                    row, col, qtw.QTableWidgetItem(str(dicti[key])))

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

        self.load_data_in_tabel(self.ui.current_stock, self.stock_sheet, len(
            self.stock_sheet), self.schema["stock_sheet"]["columns"])
        self.load_data_in_tabel(self.ui.current_stock_find_entry, [], 1, self.schema["stock_sheet"]["columns"])
        self.load_data_in_tabel(self.ui.find_stock_preview, [], 1, self.schema["stock_sheet"]["columns"])
        self.load_data_in_tabel(self.ui.current_entry_add_stock, [], 1, self.schema["stock_sheet"]["columns"])

        self.ui.tabWidget.setTabEnabled(3, True)
        self.ui.tabWidget.setTabEnabled(2, True)
        self.ui.tabWidget.setTabEnabled(1, True)
        self.ui.tabWidget.setTabEnabled(0, True)
        self.ui.actionSave_As.setEnabled(True)
        self.ui.actionSave.setEnabled(True)

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
            os.exit_(0)

    def save_warning(self, item):
        if item.text() == "Save":
            self.save_project(self.project_path)


if __name__ == "__main__":

    app = qtw.QApplication([])

    window = MainWindow()
    window.show()

    app.exec()
